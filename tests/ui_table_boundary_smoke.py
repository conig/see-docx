#!/usr/bin/env python3
"""Keep a real rightmost-column drag out of its open-layout neighbour."""

from __future__ import annotations

from pathlib import Path
import subprocess
import tempfile
import time
import zipfile
from xml.etree import ElementTree

import gi

gi.require_version("Gdk", "3.0")
gi.require_version("Gtk", "3.0")

from gi.repository import Gdk, Gio, GLib, Gtk

from see_docx.viewer import (
    APPLICATION_ID,
    DocxWindow,
    _RICH_CLIPBOARD_EMBED_SOURCE_TARGET,
    _fold_selection_whitespace,
)
from sway_test_support import (
    DesktopInput,
    WORKSPACE,
    capture_failure_screenshot,
    focus_client,
    focus_workspace,
    smoke_client,
)


TIMEOUT_MS = 120_000
SMOKE_TEST_CLASS = "codex-smoke-test"
WORDPROCESSINGML = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
TABLE = [
    [
        "Outcome indicator",
        "Contribution pathway for the work",
        "Evaluation description and expected result",
    ],
    [
        "Services respond sooner to changing local needs",
        "The programme combines regular review meetings, a secure shared dashboard, "
        "staff training and locally agreed actions so teams can identify new concerns "
        "and respond through their ordinary planning processes.",
        "Evidence will include activation records, meeting logs, response times, "
        "training completion, staff feedback, operating costs, unintended effects "
        "and a documented implementation package for other services.",
    ],
    [
        "Practical tools become part of routine delivery",
        "Teams introduce the new tool through existing governance processes and use "
        "it during ordinary service reviews, with nominated staff responsible for "
        "maintenance, interpretation and follow-up.",
        "Every participating service is activated; local coordinators are appointed; "
        "staff demonstrate the workflow; usage, reliability, timeliness, workload, "
        "action logs and improvement cycles are reported.",
    ],
    [
        "Staff adopt useful practices more rapidly",
        "Locally reviewed information is connected to approved guidance, practical "
        "resources and a clear response pathway so that emerging issues can be "
        "addressed without waiting for a distant annual report.",
        "Review-to-action pathways are mapped; response decisions are documented "
        "when signals occur; confidence, usefulness and burden are assessed; examples "
        "of changed practice are independently reviewed.",
    ],
    [
        "Communities influence design and governance",
        "Community members, service users and local leaders guide the design, "
        "governance, privacy settings and interpretation of the project through paid "
        "workshops and recurring advisory meetings.",
        "Decision records show how advice changed delivery; comprehension and "
        "usability thresholds are met; participation, acceptability, cultural safety "
        "and intention to continue are assessed and returned to the community.",
    ],
    [
        "Benefits can continue after the initial project",
        "Partner organisations agree responsibilities for oversight, maintenance, "
        "workforce development and future financing, while reusable guidance supports "
        "careful adoption in additional locations.",
        "A sustainability plan, costed extension model, technical documentation, "
        "maintenance schedule, policy brief, presentations and future funding pathway "
        "are completed and endorsed by project partners.",
    ],
]


def _open_asymmetric_table(document_xml: bytes) -> bytes:
    """Create the fixed 20/40/40 grid that Poppler reads column-major."""

    document = ElementTree.fromstring(document_xml)
    table = document.findall(f".//{WORDPROCESSINGML}tbl")[-1]
    properties = table.find(f"{WORDPROCESSINGML}tblPr")
    if properties is None:
        raise RuntimeError("generated regression table has no properties")
    properties.clear()

    def add_property(name: str, **attributes: str) -> ElementTree.Element:
        element = ElementTree.SubElement(properties, f"{WORDPROCESSINGML}{name}")
        for key, value in attributes.items():
            element.set(f"{WORDPROCESSINGML}{key}", value)
        return element

    add_property("tblW", w="9638", type="dxa")
    add_property("jc", val="start")
    add_property("tblInd", w="0", type="dxa")
    add_property("tblLayout", type="fixed")
    margins = add_property("tblCellMar")
    for name, width in (
        ("top", "0"),
        ("start", "108"),
        ("bottom", "0"),
        ("end", "108"),
    ):
        margin = ElementTree.SubElement(margins, f"{WORDPROCESSINGML}{name}")
        margin.set(f"{WORDPROCESSINGML}w", width)
        margin.set(f"{WORDPROCESSINGML}type", "dxa")
    borders = add_property("tblBorders")
    for name in ("top", "bottom"):
        border = ElementTree.SubElement(borders, f"{WORDPROCESSINGML}{name}")
        border.set(f"{WORDPROCESSINGML}val", "single")
        border.set(f"{WORDPROCESSINGML}sz", "8")
        border.set(f"{WORDPROCESSINGML}color", "000000")
    for name in ("left", "right", "insideH", "insideV"):
        border = ElementTree.SubElement(borders, f"{WORDPROCESSINGML}{name}")
        border.set(f"{WORDPROCESSINGML}val", "nil")

    widths = ("1910", "3861", "3867")
    grid = table.find(f"{WORDPROCESSINGML}tblGrid")
    if grid is None or len(grid) != len(widths):
        raise RuntimeError("generated regression table has an unexpected grid")
    for column, width in zip(grid, widths, strict=True):
        column.set(f"{WORDPROCESSINGML}w", width)
    for row in table.findall(f"{WORDPROCESSINGML}tr"):
        for cell, width in zip(
            row.findall(f"{WORDPROCESSINGML}tc"), widths, strict=True
        ):
            cell_properties = cell.find(f"{WORDPROCESSINGML}tcPr")
            if cell_properties is None:
                raise RuntimeError("generated regression cell has no properties")
            cell_properties.clear()
            cell_width = ElementTree.SubElement(
                cell_properties, f"{WORDPROCESSINGML}tcW"
            )
            cell_width.set(f"{WORDPROCESSINGML}w", width)
            cell_width.set(f"{WORDPROCESSINGML}type", "dxa")
            cell_borders = ElementTree.SubElement(
                cell_properties, f"{WORDPROCESSINGML}tcBorders"
            )
            for name in ("top", "left", "bottom", "right"):
                border = ElementTree.SubElement(
                    cell_borders, f"{WORDPROCESSINGML}{name}"
                )
                border.set(f"{WORDPROCESSINGML}val", "nil")
            vertical_alignment = ElementTree.SubElement(
                cell_properties, f"{WORDPROCESSINGML}vAlign"
            )
            vertical_alignment.set(f"{WORDPROCESSINGML}val", "center")
    # Offset the short first body cell within its much taller row. LibreOffice
    # then emits its glyphs between lines from the neighbouring cell, matching
    # the column-interleaved Poppler order that exposed the double-click bug.
    first_body_cell = table.findall(f"{WORDPROCESSINGML}tr")[1].findall(
        f"{WORDPROCESSINGML}tc"
    )[0]
    paragraph_properties = first_body_cell.find(
        f"{WORDPROCESSINGML}p/{WORDPROCESSINGML}pPr"
    )
    if paragraph_properties is None:
        raise RuntimeError("generated first body cell has no paragraph properties")
    spacing = paragraph_properties.find(f"{WORDPROCESSINGML}spacing")
    if spacing is None:
        spacing = ElementTree.SubElement(
            paragraph_properties, f"{WORDPROCESSINGML}spacing"
        )
    spacing.set(f"{WORDPROCESSINGML}before", "480")
    spacing.set(f"{WORDPROCESSINGML}after", "0")
    return ElementTree.tostring(document, encoding="utf-8", xml_declaration=True)


def _fixture(destination: Path) -> None:
    markdown = destination.with_suffix(".md")
    markdown.write_text(
        "\n".join(
            (
                "| " + " | ".join(TABLE[0]) + " |",
                "|" + "|".join(["---"] * 3) + "|",
                *("| " + " | ".join(row) + " |" for row in TABLE[1:]),
            )
        ),
        encoding="utf-8",
    )
    subprocess.run(
        ["pandoc", str(markdown), "-o", str(destination)],
        check=True,
        capture_output=True,
        text=True,
    )
    with zipfile.ZipFile(destination) as archive:
        entries = {
            info.filename: (info, archive.read(info.filename))
            for info in archive.infolist()
        }
    document_info, document_xml = entries["word/document.xml"]
    entries["word/document.xml"] = (
        document_info,
        _open_asymmetric_table(document_xml),
    )
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, (info, data) in entries.items():
            archive.writestr(info, data)


def _cell_character_index(
    page_text: str, cell_text: str, *, inside: bool
) -> int | None:
    rendered, rendered_indices = _fold_selection_whitespace(page_text)
    cell, _cell_indices = _fold_selection_whitespace(cell_text)
    needle = cell[:8]
    start = rendered.find(needle)
    if not needle or start < 0 or rendered.find(needle, start + 1) >= 0:
        return None
    return rendered_indices[start + min(4, len(needle) - 1) if inside else start]


def main() -> int:
    GLib.set_prgname(SMOKE_TEST_CLASS)
    with tempfile.TemporaryDirectory(
        prefix="see-docx-table-boundary-smoke-"
    ) as directory:
        fixture = Path(directory) / "open-asymmetric-table.docx"
        _fixture(fixture)
        result: dict[str, str] = {}
        state: dict[str, object] = {"phase": "double_click_first"}
        desktop_input: DesktopInput | None = None
        application = Gtk.Application(
            application_id=f"{APPLICATION_ID}.table-boundary-smoke",
            flags=Gio.ApplicationFlags.NON_UNIQUE,
        )

        def finish(window: DocxWindow, message: str | None = None) -> bool:
            nonlocal desktop_input
            if message:
                result["error"] = message
                capture_failure_screenshot()
            if desktop_input is not None:
                desktop_input.close()
                desktop_input = None
            window.close()
            application.quit()
            return GLib.SOURCE_REMOVE

        def check(window: DocxWindow) -> bool:
            nonlocal desktop_input
            try:
                if not window.document.has_document or not window.document._pages:
                    return GLib.SOURCE_CONTINUE
                client = smoke_client(SMOKE_TEST_CLASS)
                if client is None or client["workspace"]["id"] != int(WORKSPACE):
                    return GLib.SOURCE_CONTINUE
                source = window.document._rich_source
                if source is None:
                    return GLib.SOURCE_CONTINUE
                table_index = max(
                    fragment.table_index
                    for fragment in source._fragments
                    if fragment.table_index is not None
                )
                fragments = {
                    (fragment.row_index, fragment.column_start): fragment
                    for fragment in source._fragments
                    if fragment.table_index == table_index
                }
                final_column = max(
                    column for _row, column in fragments if column is not None
                )
                locations: list[tuple[int, object, int, tuple[float, float]]] = []
                for (row, column), fragment in fragments.items():
                    if row is None or column != final_column:
                        continue
                    for page in window.document._pages:
                        index = _cell_character_index(
                            page._page.get_text(), fragment.text, inside=True
                        )
                        if index is None:
                            continue
                        has_layout, rectangles = page._page.get_text_layout()
                        if not has_layout:
                            continue
                        rectangle = rectangles[index]
                        locations.append(
                            (
                                row,
                                page,
                                index,
                                (
                                    (float(rectangle.x1) + float(rectangle.x2))
                                    * page._zoom
                                    / 2,
                                    (float(rectangle.y1) + float(rectangle.y2))
                                    * page._zoom
                                    / 2,
                                ),
                            )
                        )
                        break
                candidates = [
                    (lower[0] - upper[0], upper, lower)
                    for upper in locations
                    for lower in locations
                    if upper[0] < lower[0] and upper[1] is lower[1]
                ]
                if not candidates:
                    return GLib.SOURCE_CONTINUE
                _distance, upper, lower = max(candidates, key=lambda item: item[0])
                page = upper[1]

                if desktop_input is None:
                    desktop_input = DesktopInput(directory)
                    focus_workspace(WORKSPACE)
                    focus_client(client)
                    page_allocation = page.get_allocation()
                    midpoint = (upper[3][1] + lower[3][1]) / 2
                    adjustment = window.document.widget.get_vadjustment()
                    target = (
                        page_allocation.y
                        + midpoint
                        - adjustment.get_page_size() / 2
                    )
                    adjustment.set_value(
                        min(
                            max(target, adjustment.get_lower()),
                            adjustment.get_lower()
                            + window.document._maximum_scroll(),
                        )
                    )
                    state["selection"] = (upper, lower, fragments, final_column)
                    state["pointer_ready_at"] = time.monotonic() + 0.7
                    return GLib.SOURCE_CONTINUE
                if time.monotonic() < float(state["pointer_ready_at"]):
                    return GLib.SOURCE_CONTINUE

                upper, lower, fragments, final_column = state["selection"]
                page = upper[1]
                # The reported failure is at the leading word of the second
                # cell in the first body row.  Its glyphs can precede the
                # vertically offset first cell in Poppler's text order, but a
                # real double-click must still select this cell alone rather
                # than the header cell directly above it.
                double_clicked_cell = fragments[(1, 1)]
                double_clicked_cell_index = _cell_character_index(
                    page._page.get_text(), double_clicked_cell.text, inside=False
                )
                if double_clicked_cell_index is None:
                    return finish(
                        window,
                        "could not locate the second body-column cell for double-click",
                    )
                has_layout, page_rectangles = page._page.get_text_layout()
                if not has_layout:
                    return finish(window, "the rendered table has no glyph layout")
                double_clicked_rectangle = page_rectangles[double_clicked_cell_index]
                double_clicked_point = (
                    (
                        float(double_clicked_rectangle.x1)
                        + float(double_clicked_rectangle.x2)
                    )
                    * page._zoom
                    / 2,
                    (
                        float(double_clicked_rectangle.y1)
                        + float(double_clicked_rectangle.y2)
                    )
                    * page._zoom
                    / 2,
                )
                translated = page.translate_coordinates(window, 0, 0)
                page_x, page_y = translated[-2:]
                window_x, window_y = client["at"]
                double_clicked_x = window_x + page_x + double_clicked_point[0]
                double_clicked_y = window_y + page_y + double_clicked_point[1]

                if state["phase"] == "double_click_first":
                    desktop_input.move_cursor(
                        int(double_clicked_x), int(double_clicked_y)
                    )
                    desktop_input.left_button(pressed=True)
                    desktop_input.left_button(pressed=False)
                    state["phase"] = "double_click_second"
                    state["second_click_at"] = time.monotonic() + 0.12
                    return GLib.SOURCE_CONTINUE
                if state["phase"] == "double_click_second":
                    if time.monotonic() < float(state["second_click_at"]):
                        return GLib.SOURCE_CONTINUE
                    desktop_input.move_cursor(
                        int(double_clicked_x), int(double_clicked_y)
                    )
                    desktop_input.left_button(pressed=True)
                    desktop_input.left_button(pressed=False)
                    state["phase"] = "double_click_check"
                    state["double_click_check_at"] = time.monotonic() + 0.3
                    return GLib.SOURCE_CONTINUE
                if state["phase"] == "double_click_check":
                    if time.monotonic() < float(state["double_click_check_at"]):
                        return GLib.SOURCE_CONTINUE
                    clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)
                    actual = clipboard.wait_for_text()
                    html_data = clipboard.wait_for_contents(
                        Gdk.Atom.intern("text/html", False)
                    )
                    odt_data = clipboard.wait_for_contents(
                        Gdk.Atom.intern(
                            _RICH_CLIPBOARD_EMBED_SOURCE_TARGET.decode(), False
                        )
                    )
                    rendered_selection = (
                        page._layout_text_selection(page._text_selection)
                        if page._text_selection is not None
                        else None
                    )
                    selected_rectangles = (
                        rendered_selection[1]
                        if rendered_selection is not None
                        else []
                    )
                    selected_bounds = {
                        (
                            round(float(rectangle.x1), 3),
                            round(float(rectangle.y1), 3),
                            round(float(rectangle.x2), 3),
                            round(float(rectangle.y2), 3),
                        )
                        for rectangle in selected_rectangles
                    }
                    target_top = min(
                        float(double_clicked_rectangle.y1),
                        float(double_clicked_rectangle.y2),
                    )
                    leaking_above = [
                        rectangle
                        for rectangle in selected_rectangles
                        if (float(rectangle.y1) + float(rectangle.y2)) / 2
                        < target_top - 0.5
                    ]
                    page_text = page._page.get_text()
                    missing_markers: list[str] = []
                    for marker in (
                        "programme",
                        "combines",
                        "regular",
                        "secure",
                        "agreed",
                        "concerns",
                    ):
                        marker_start = page_text.find(marker)
                        if (
                            marker_start < 0
                            or page_text.find(marker, marker_start + 1) >= 0
                        ):
                            missing_markers.append(f"{marker} (unlocatable)")
                            continue
                        for index in range(
                            marker_start, marker_start + len(marker)
                        ):
                            rectangle = page_rectangles[index]
                            bounds = (
                                round(float(rectangle.x1), 3),
                                round(float(rectangle.y1), 3),
                                round(float(rectangle.x2), 3),
                                round(float(rectangle.y2), 3),
                            )
                            if bounds not in selected_bounds:
                                missing_markers.append(marker)
                                break
                    if (
                        actual != double_clicked_cell.text
                        or html_data is None
                        or odt_data is None
                        or missing_markers
                        or leaking_above
                    ):
                        return finish(
                            window,
                            "double-clicking the second cell in the first body row "
                            "did not visibly select only that cell: "
                            f"expected={double_clicked_cell.text!r}; actual={actual!r}; "
                            f"selected_glyphs={len(selected_rectangles)}; "
                            f"missing_markers={missing_markers!r}; "
                            f"leaking_above={len(leaking_above)}; "
                            f"has_html={html_data is not None}; "
                            f"has_odf={odt_data is not None}",
                        )
                    state["phase"] = "locate"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "locate":
                    origin_x, origin_y = window_x + page_x, window_y + page_y
                    start_x, start_y = lower[3]
                    end_x, end_y = upper[3]
                    desktop_input.move_cursor(
                        int(origin_x + start_x), int(origin_y + start_y)
                    )
                    desktop_input.left_button(pressed=True)
                    desktop_input.move_cursor(
                        int(origin_x + end_x), int(origin_y + end_y)
                    )
                    desktop_input.left_button(pressed=False)
                    state["phase"] = "check"
                    state["check_at"] = time.monotonic() + 0.3
                    return GLib.SOURCE_CONTINUE
                if time.monotonic() < float(state["check_at"]):
                    return GLib.SOURCE_CONTINUE

                first_row, last_row = upper[0], lower[0]
                expected = "\n".join(
                    fragments[(row, final_column)].text
                    for row in range(first_row, last_row + 1)
                )
                clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)
                actual = clipboard.wait_for_text()
                html_data = clipboard.wait_for_contents(
                    Gdk.Atom.intern("text/html", False)
                )
                odt_data = clipboard.wait_for_contents(
                    Gdk.Atom.intern(
                        _RICH_CLIPBOARD_EMBED_SOURCE_TARGET.decode(), False
                    )
                )
                html_tables = (
                    bytes(html_data.get_data()).decode("utf-8").count("<table")
                    if html_data is not None
                    else 0
                )
                if actual != expected or html_tables != 1 or odt_data is None:
                    return finish(
                        window,
                        "the synthetic final-column drag copied the wrong grid: "
                        f"expected_length={len(expected)}; "
                        f"actual_length={len(actual or '')}; "
                        f"html_tables={html_tables}; has_odf={odt_data is not None}",
                    )

                rendered_selection = page._layout_text_selection(
                    page._text_selection
                )
                if rendered_selection is None:
                    return finish(window, "the final column has no visible highlight")
                _selected_text, selected_rectangles = rendered_selection
                page_text = page._page.get_text()
                has_layout, page_rectangles = page._page.get_text_layout()
                confirmed_starts = [
                    _cell_character_index(
                        page_text,
                        fragments[(row, final_column)].text,
                        inside=False,
                    )
                    for row in range(first_row, last_row + 1)
                ]
                if not has_layout or any(
                    index is None for index in confirmed_starts
                ):
                    return finish(
                        window,
                        "could not establish the rendered final-column boundary",
                    )
                expected_left = min(
                    float(page_rectangles[index].x1)
                    for index in confirmed_starts
                    if index is not None
                )
                leaking_rectangles = [
                    rectangle
                    for rectangle in selected_rectangles
                    if (float(rectangle.x1) + float(rectangle.x2)) / 2
                    < expected_left
                ]
                if leaking_rectangles:
                    return finish(
                        window,
                        "the visible final-column selection spilled into its "
                        "wide neighbour: "
                        f"selected_glyphs={len(selected_rectangles)}; "
                        f"leaking_glyphs={len(leaking_rectangles)}; "
                        f"expected_left={expected_left:.2f}; "
                        "leak_left="
                        f"{min(float(rectangle.x1) for rectangle in leaking_rectangles):.2f}",
                    )
                print(
                    "an open asymmetric table visibly selects an interleaved "
                    "double-clicked cell and keeps its final-column highlight "
                    "out of the neighbouring column"
                )
                return finish(window)
            except (OSError, RuntimeError, subprocess.SubprocessError) as error:
                return finish(window, str(error))

        def activate(app: Gtk.Application) -> None:
            window = DocxWindow(app, fixture)
            window.set_default_size(1440, 900)
            window.show_all()
            GLib.timeout_add(100, check, window)
            GLib.timeout_add(
                TIMEOUT_MS,
                finish,
                window,
                "timed out verifying the open-table final-column boundary",
            )

        application.connect("activate", activate)
        application.run(["see-docx-table-boundary-smoke"])
    if error := result.get("error"):
        raise RuntimeError(error)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
