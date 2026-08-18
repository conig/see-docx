#!/usr/bin/env python3
"""Reproduce table hit-testing after a long, XML-divergent DOCX prelude."""

from __future__ import annotations

from pathlib import Path
import subprocess
import tempfile
import time
import zipfile

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
FILLER = (
    "Ordinary repeated project information provides background context for "
    "the document and continues across the page."
)
TABLE = [
    ["Measure of success", "How the work contributes", "Description of the outcome"],
    [
        "Evidence informs action",
        "The project produces timely information for local decisions and shared learning.",
        "A useful report is delivered and reviewed by the group.",
    ],
    [
        "New tools are adopted",
        "Teams use the practical tool in routine planning and review.",
        "The tool is active at each site and staff can use it.",
    ],
    [
        "Practice changes sooner",
        "Locally reviewed information is connected to a clear response pathway.",
        "Review and response steps are documented and used.",
    ],
    [
        "Community participation grows",
        "Community members guide design interpretation and communication.",
        "Decision records show how community advice changed delivery.",
    ],
    [
        "Skills and confidence increase",
        "Training and mentoring build local capability throughout delivery.",
        "Participants report confidence and demonstrate the workflow.",
    ],
    [
        "Benefits continue over time",
        "Governance and maintenance plans support use after the pilot.",
        "Partners agree responsibilities and a sustainable review cycle.",
    ],
]


def _fixture(destination: Path) -> None:
    """Create a long DOCX whose repeated running header diverges from body XML."""

    markdown = destination.with_suffix(".md")
    lines = ["\n\n".join([FILLER] * 260), ""]
    lines.extend(
        (
            "| " + " | ".join(TABLE[0]) + " |",
            "|" + "|".join(["---"] * 3) + "|",
            *("| " + " | ".join(row) + " |" for row in TABLE[1:]),
        )
    )
    markdown.write_text("\n".join(lines), encoding="utf-8")
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
    document_info, document_data = entries["word/document.xml"]
    document_xml = document_data.decode("utf-8")
    section = document_xml.index("<w:sectPr")
    insertion = document_xml.index(">", section) + 1
    document_xml = (
        document_xml[:insertion]
        + '<w:headerReference w:type="default" r:id="rIdRunningHeader"/>'
        + document_xml[insertion:]
    )
    entries["word/document.xml"] = (
        document_info,
        document_xml.encode("utf-8"),
    )

    relationships_info, relationships_data = entries[
        "word/_rels/document.xml.rels"
    ]
    relationships_xml = relationships_data.decode("utf-8").replace(
        "</Relationships>",
        '<Relationship Id="rIdRunningHeader" '
        'Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" '
        'Target="header1.xml"/></Relationships>',
    )
    entries["word/_rels/document.xml.rels"] = (
        relationships_info,
        relationships_xml.encode("utf-8"),
    )

    content_types_info, content_types_data = entries["[Content_Types].xml"]
    content_types_xml = content_types_data.decode("utf-8").replace(
        "</Types>",
        '<Override PartName="/word/header1.xml" '
        'ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>'
        "</Types>",
    )
    entries["[Content_Types].xml"] = (
        content_types_info,
        content_types_xml.encode("utf-8"),
    )
    header = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:p><w:r><w:t>REPEATED RUNNING HEADER FOR EVERY PAGE</w:t></w:r></w:p>"
        "</w:hdr>"
    ).encode("utf-8")

    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, (info, data) in entries.items():
            archive.writestr(info, data)
        archive.writestr("word/header1.xml", header)


def _cell_prefix_index(page_text: str, cell_text: str) -> int | None:
    """Locate a short unique cell prefix despite PDF line-wrap differences."""

    rendered, rendered_indices = _fold_selection_whitespace(page_text)
    cell, _cell_indices = _fold_selection_whitespace(cell_text)
    needle = cell[:8]
    start = rendered.find(needle)
    if not needle or start < 0 or rendered.find(needle, start + 1) >= 0:
        return None
    return rendered_indices[start + min(4, len(needle) - 1)]


def _cell_start_index(page_text: str, cell_text: str) -> int | None:
    """Locate the first glyph of a uniquely identifiable rendered cell."""

    rendered, rendered_indices = _fold_selection_whitespace(page_text)
    cell, _cell_indices = _fold_selection_whitespace(cell_text)
    needle = cell[:8]
    start = rendered.find(needle)
    if not needle or start < 0 or rendered.find(needle, start + 1) >= 0:
        return None
    return rendered_indices[start]


def main() -> int:
    GLib.set_prgname(SMOKE_TEST_CLASS)
    with tempfile.TemporaryDirectory(prefix="see-docx-table-mapping-smoke-") as directory:
        fixture = Path(directory) / "long-table.docx"
        _fixture(fixture)
        result: dict[str, str] = {}
        state: dict[str, object] = {"phase": "locate_first"}
        desktop_input: DesktopInput | None = None
        application = Gtk.Application(
            application_id=f"{APPLICATION_ID}.table-mapping-smoke",
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
                if not window.document.has_document:
                    return GLib.SOURCE_CONTINUE
                client = smoke_client(SMOKE_TEST_CLASS)
                if client is None or client["workspace"]["id"] != int(WORKSPACE):
                    return GLib.SOURCE_CONTINUE
                source = window.document._rich_source
                if source is None or window.document._maximum_scroll() <= 0:
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
                last_column = max(
                    column
                    for row, column in fragments
                    if row is not None and column is not None
                )
                target_column = (
                    last_column
                    if state["phase"] in {"locate_last", "check_last"}
                    else 0
                )
                locations: list[tuple[int, int, object, tuple[float, float]]] = []
                for (row, column), fragment in fragments.items():
                    if column != target_column or row is None:
                        continue
                    for page_index, page in enumerate(window.document._pages):
                        rendered_index = _cell_prefix_index(
                            page._page.get_text(), fragment.text
                        )
                        if rendered_index is None:
                            continue
                        has_layout, rectangles = page._page.get_text_layout()
                        if not has_layout:
                            continue
                        rectangle = rectangles[rendered_index]
                        locations.append(
                            (
                                row,
                                page_index,
                                page,
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
                    if upper[0] < lower[0] and upper[1] == lower[1]
                ]
                if not candidates:
                    return GLib.SOURCE_CONTINUE
                _distance, upper, lower = max(candidates, key=lambda item: item[0])

                if desktop_input is None:
                    desktop_input = DesktopInput(directory)
                    focus_workspace(WORKSPACE)
                    focus_client(client)
                    page = upper[2]
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
                            adjustment.get_lower() + window.document._maximum_scroll(),
                        )
                    )
                    state["selection"] = (
                        upper,
                        lower,
                        fragments,
                        target_column,
                    )
                    state["pointer_ready_at"] = time.monotonic() + 0.7
                    return GLib.SOURCE_CONTINUE
                if time.monotonic() < float(state["pointer_ready_at"]):
                    return GLib.SOURCE_CONTINUE

                if state["phase"] in {"locate_first", "locate_last"}:
                    state["selection"] = (
                        upper,
                        lower,
                        fragments,
                        target_column,
                    )
                    page = upper[2]
                    translated = page.translate_coordinates(window, 0, 0)
                    page_x, page_y = translated[-2:]
                    window_x, window_y = client["at"]
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
                    state["phase"] = (
                        "check_last"
                        if state["phase"] == "locate_last"
                        else "check_first"
                    )
                    state["check_at"] = time.monotonic() + 0.3
                    return GLib.SOURCE_CONTINUE
                if time.monotonic() < float(state["check_at"]):
                    return GLib.SOURCE_CONTINUE

                upper, lower, fragments, target_column = state["selection"]
                page = upper[2]
                first_row, last_row = upper[0], lower[0]
                expected = "\n".join(
                    fragments[(row, target_column)].text
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
                        "an upward drag in the long DOCX spilled out of its column: "
                        f"rows={first_row}-{last_row}; "
                        f"expected_length={len(expected)}; "
                        f"actual_length={len(actual or '')}; "
                        f"html_tables={html_tables}; has_odf={odt_data is not None}",
                    )
                if state["phase"] == "check_first":
                    state["phase"] = "locate_last"
                    return GLib.SOURCE_CONTINUE

                rendered_selection = page._layout_text_selection(
                    page._text_selection
                )
                if rendered_selection is None:
                    return finish(
                        window,
                        "the rightmost table column produced no visible selection",
                    )
                _selected_text, selected_rectangles = rendered_selection
                page_text = page._page.get_text()
                has_layout, page_rectangles = page._page.get_text_layout()
                confirmed_starts = [
                    _cell_start_index(
                        page_text,
                        fragments[(row, target_column)].text,
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
                # This is intentionally independent of See DOCX's inferred
                # cell bands: using those bands here previously let the test
                # validate the same wrong boundary that painted the spill.
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
                        "the visible rightmost-column selection spilled into "
                        "the adjacent column: "
                        f"selected_glyphs={len(selected_rectangles)}; "
                        f"leaking_glyphs={len(leaking_rectangles)}; "
                        f"expected_left={expected_left:.2f}; "
                        "leak_left="
                        f"{min(float(rectangle.x1) for rectangle in leaking_rectangles):.2f}",
                    )
                print(
                    "long XML-divergent DOCX keeps upward drags inside the first "
                    "and final table columns"
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
                "timed out verifying long-document table mapping",
            )

        application.connect("activate", activate)
        application.run(["see-docx-table-mapping-smoke"])
    if error := result.get("error"):
        raise RuntimeError(error)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
