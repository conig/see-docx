#!/usr/bin/env python3
"""Verify table-shaped GTK selections and Writer row/column replacement."""

from __future__ import annotations

import argparse
from io import BytesIO
import json
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
GRID = [
    [
        "A1 outer-left baseline",
        "B1 centre first has a wrapped phrase for realistic table content",
        "C1 outer-right baseline",
    ],
    [
        "A2 outer-left middle",
        "B2 centre selected words inside a realistically wrapped cell",
        "C2 outer-right middle",
    ],
    [
        "A3 outer-left final",
        "B3 centre last has another wrapped phrase for realistic table content",
        "C3 outer-right final contents",
    ],
]
COMPLETE_GRID = [["A", "B", "C"], *GRID]


def _fixture(destination: Path) -> None:
    source = Path(__file__).with_name("fixtures") / "rich_selection_grid.md"
    subprocess.run(
        ["pandoc", str(source), "-o", str(destination)],
        check=True,
        capture_output=True,
        text=True,
    )


def _glyph_character_point(
    page: object, target: str, *, last: bool = False
) -> tuple[float, float]:
    """Return the centre of the first or last glyph in one unique marker."""

    text = page._page.get_text()
    index = text.find(target)
    has_layout, rectangles = page._page.get_text_layout()
    if index < 0 or text.find(target, index + 1) >= 0 or not has_layout:
        raise RuntimeError(f"rendered PDF has no unique glyph marker for {target!r}")
    rectangle = rectangles[index + len(target) - 1 if last else index]
    return (
        (float(rectangle.x1) + float(rectangle.x2)) * page._zoom / 2,
        (float(rectangle.y1) + float(rectangle.y2)) * page._zoom / 2,
    )


def _drag_characters(
    page: object,
    client: dict[str, object],
    desktop_input: DesktopInput,
    start_text: str,
    end_text: str,
) -> None:
    """Drag from the first glyph of start_text to the last of end_text."""

    start_x, start_y = _glyph_character_point(page, start_text)
    end_x, end_y = _glyph_character_point(page, end_text, last=True)
    window_x, window_y = client["at"]
    page_local_x, page_local_y = page.get_window().get_root_coords(0, 0)
    origin_x = window_x + page_local_x
    origin_y = window_y + page_local_y
    desktop_input.move_cursor(int(origin_x + start_x), int(origin_y + start_y))
    desktop_input.left_button(pressed=True)
    desktop_input.move_cursor(int(origin_x + end_x), int(origin_y + end_y))
    desktop_input.left_button(pressed=False)


def _click(
    page: object,
    client: dict[str, object],
    desktop_input: DesktopInput,
    target: str,
) -> None:
    x, y = _glyph_character_point(page, target)
    window_x, window_y = client["at"]
    page_local_x, page_local_y = page.get_window().get_root_coords(0, 0)
    desktop_input.move_cursor(
        int(window_x + page_local_x + x), int(window_y + page_local_y + y)
    )
    desktop_input.left_button(pressed=True)
    desktop_input.left_button(pressed=False)


def _clipboard_odt(*, text: str, rows: int, cells: int) -> bytes:
    clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)
    actual_text = clipboard.wait_for_text()
    html_data = clipboard.wait_for_contents(Gdk.Atom.intern("text/html", False))
    odt_data = clipboard.wait_for_contents(
        Gdk.Atom.intern(_RICH_CLIPBOARD_EMBED_SOURCE_TARGET.decode(), False)
    )
    if actual_text != text or html_data is None or odt_data is None:
        raise RuntimeError(
            "table selection did not offer matching text, HTML, and embedded ODF: "
            f"expected_text={text!r}; actual_text={actual_text!r}; "
            f"has_html={html_data is not None}; has_odf={odt_data is not None}"
        )
    html = bytes(html_data.get_data()).decode()
    if (
        html.count("<table") != 1
        or html.count("<tr>") != rows
        or html.count("<td") != cells
    ):
        raise RuntimeError("table selection was not exported as one HTML cell grid")
    odt = bytes(odt_data.get_data())
    with zipfile.ZipFile(BytesIO(odt)) as archive:
        content = archive.read("content.xml").decode()
    if (
        content.count("<table:table-row>") != rows
        or content.count("<table:table-cell") != cells
    ):
        raise RuntimeError("embedded Writer table has the wrong selected-cell grid")
    return odt


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument(
        "--case",
        choices=(
            "all",
            "table-copy",
            "single-cell",
            "double-click",
            "column",
            "paste",
        ),
        default="all",
        help="run one focused stage or the complete integration flow",
    )
    arguments = parser.parse_args()
    GLib.set_prgname(SMOKE_TEST_CLASS)
    with tempfile.TemporaryDirectory(
        prefix="see-docx-rich-selection-smoke-"
    ) as directory:
        fixture = Path(directory) / "grid.docx"
        _fixture(fixture)
        result: dict[str, str] = {}
        initial_phases = {
            "all": "table_hover",
            "table-copy": "table_hover",
            "single-cell": "single_cell",
            "double-click": "double_click_first",
            "column": "column",
            "paste": "row",
        }
        state: dict[str, object] = {"phase": initial_phases[arguments.case]}
        probe: subprocess.Popen[str] | None = None
        desktop_input: DesktopInput | None = None
        application = Gtk.Application(
            application_id=f"{APPLICATION_ID}.rich-selection-smoke",
            flags=Gio.ApplicationFlags.NON_UNIQUE,
        )

        def finish(window: DocxWindow, message: str | None = None) -> bool:
            nonlocal desktop_input, probe
            if message:
                result["error"] = message
                capture_failure_screenshot()
            if probe is not None and probe.poll() is None:
                probe.terminate()
                try:
                    probe.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    probe.kill()
                    probe.wait(timeout=3)
            probe = None
            if desktop_input is not None:
                desktop_input.close()
                desktop_input = None
            window.close()
            application.quit()
            return GLib.SOURCE_REMOVE

        def start_probe(mode: str) -> None:
            nonlocal probe
            probe = subprocess.Popen(
                [
                    "python3",
                    str(Path(__file__).with_name("libreoffice_table_paste_probe.py")),
                    mode,
                ],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )

        def completed_probe(
            expected_selection: list[list[str]],
        ) -> tuple[bool, str | None]:
            nonlocal probe
            if probe is None or probe.poll() is None:
                return False, None
            stdout, stderr = probe.communicate()
            return_code = probe.returncode
            probe = None
            if return_code:
                return True, f"LibreOffice paste probe failed: {stderr.strip()}"
            try:
                actual = json.loads(stdout)
            except json.JSONDecodeError as error:
                return True, f"LibreOffice paste probe returned invalid JSON: {error}"
            if actual.get("blank_tables") != [expected_selection]:
                return True, (
                    "pasting See DOCX's live clipboard into a blank Writer "
                    "document did not create exactly one coherent table: "
                    f"expected={[expected_selection]!r}, "
                    f"actual={actual.get('blank_tables')!r}, "
                    f"plain_text={actual.get('blank_text')!r}"
                )
            return True, None

        def check(window: DocxWindow) -> bool:
            nonlocal desktop_input
            try:
                if not window.document.has_document or not window.document._pages:
                    return GLib.SOURCE_CONTINUE
                page = window.document._pages[0]
                if page.get_allocated_width() <= 1 or page.get_allocated_height() <= 1:
                    return GLib.SOURCE_CONTINUE
                client = smoke_client(SMOKE_TEST_CLASS)
                if client is None or client["workspace"]["id"] != int(WORKSPACE):
                    return finish(window, "rich-selection smoke did not map on workspace 15")
                if desktop_input is None:
                    desktop_input = DesktopInput(directory)
                    focus_workspace(WORKSPACE)
                    focus_client(client)
                    # Return to GTK so it can bind wl_pointer after the
                    # virtual device adds pointer capability to the seat.
                    state["pointer_ready_at"] = time.monotonic() + 0.5
                    return GLib.SOURCE_CONTINUE
                if time.monotonic() < float(state["pointer_ready_at"]):
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "table_hover":
                    x, y = _glyph_character_point(page, "B2")
                    window_x, window_y = client["at"]
                    page_x, page_y = page.get_window().get_root_coords(0, 0)
                    desktop_input.move_cursor(
                        int(window_x + page_x + x),
                        int(window_y + page_y + y),
                    )
                    state["phase"] = "table_button"
                    state["table_button_at"] = time.monotonic() + 0.2
                    state["table_hover_attempts"] = 1
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "table_button":
                    if time.monotonic() < float(state["table_button_at"]):
                        return GLib.SOURCE_CONTINUE
                    table_index = page._table_copy_hover_index
                    if table_index is None:
                        x, y = _glyph_character_point(page, "B2")
                        attempts = int(state["table_hover_attempts"])
                        if attempts < 5:
                            window_x, window_y = client["at"]
                            page_x, page_y = page.get_window().get_root_coords(0, 0)
                            desktop_input.move_cursor(
                                int(window_x + page_x + x + attempts % 2),
                                int(window_y + page_y + y),
                            )
                            state["table_hover_attempts"] = attempts + 1
                            state["table_button_at"] = time.monotonic() + 0.12
                            return GLib.SOURCE_CONTINUE
                        fragment = page.table_fragment_at(
                            (x / page._zoom, y / page._zoom)
                        )
                        return finish(
                            window,
                            "hovering a table did not reveal its copy control: "
                            f"direct_hit={None if fragment is None else (fragment.row_index, fragment.column_start)}; "
                            f"layouts={len(page._table_cell_layouts)}; "
                            f"point={(x / page._zoom, y / page._zoom)!r}; "
                            f"selection_pressed={page._selection_button_pressed}; "
                            f"copy_pressed={page._table_copy_pressed}",
                        )
                    button = page._table_copy_button(table_index)
                    if button is None:
                        return finish(window, "the table copy control has no hit box")
                    table_left = min(
                        layout.left
                        for layout in page._table_cell_layouts
                        if layout.fragment.table_index == table_index
                    ) * page._zoom
                    table_top = min(
                        layout.top
                        for layout in page._table_cell_layouts
                        if layout.fragment.table_index == table_index
                    ) * page._zoom
                    if button.right > table_left + 0.01:
                        return finish(
                            window,
                            "the table copy control obscures the table contents",
                        )
                    state["table_visible_left"] = table_left
                    state["table_visible_top"] = table_top
                    target_x = (button.left + button.right) / 2
                    target_y = (button.top + button.bottom) / 2
                    origin_x, origin_y = _glyph_character_point(page, "B2")
                    # A real user generates many intermediate motion events while
                    # moving from a cell to the adjacent control.  Teleporting the
                    # virtual pointer straight to its centre hid the regression in
                    # which the control vanished as soon as the pointer crossed a
                    # sparse part of the rendered table.
                    state["table_transition"] = [
                        (
                            origin_x + (target_x - origin_x) * step / 16,
                            origin_y + (target_y - origin_y) * step / 16,
                        )
                        for step in range(1, 17)
                    ]
                    state["table_transition_step"] = 0
                    state["phase"] = "table_transition"
                    state["table_transition_at"] = time.monotonic() + 0.04
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "table_transition":
                    if time.monotonic() < float(state["table_transition_at"]):
                        return GLib.SOURCE_CONTINUE
                    step = int(state["table_transition_step"])
                    if step and page._table_copy_hover_index is None:
                        return finish(
                            window,
                            "the copy control disappeared while the pointer moved "
                            "from the table to the control",
                        )
                    points = state["table_transition"]
                    if step < len(points):
                        target_x, target_y = points[step]
                        window_x, window_y = client["at"]
                        page_x, page_y = page.get_window().get_root_coords(0, 0)
                        desktop_input.move_cursor(
                            int(window_x + page_x + target_x),
                            int(window_y + page_y + target_y),
                        )
                        state["table_transition_step"] = step + 1
                        state["table_transition_at"] = time.monotonic() + 0.04
                        return GLib.SOURCE_CONTINUE
                    button = page._table_copy_button(
                        page._table_copy_hover_index
                    )
                    if (
                        button is None
                        or button.right
                        > float(state["table_visible_left"]) - 4.0
                        or button.bottom
                        > float(state["table_visible_top"]) - 4.0
                    ):
                        return finish(
                            window,
                            "the table copy control is painted on the table instead "
                            "of perceptibly beside it",
                        )
                    state["phase"] = "table_click"
                    state["table_click_at"] = time.monotonic() + 0.15
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "table_click":
                    if time.monotonic() < float(state["table_click_at"]):
                        return GLib.SOURCE_CONTINUE
                    if not page._table_copy_button_hot:
                        return finish(
                            window,
                            "the copy control disappeared before it could be clicked",
                        )
                    desktop_input.left_button(pressed=True)
                    desktop_input.left_button(pressed=False)
                    state["phase"] = "table_copy_check"
                    state["table_copy_check_at"] = time.monotonic() + 0.25
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "table_copy_check":
                    if time.monotonic() < float(state["table_copy_check_at"]):
                        return GLib.SOURCE_CONTINUE
                    _clipboard_odt(
                        text="\n".join("\t".join(row) for row in COMPLETE_GRID),
                        rows=len(COMPLETE_GRID),
                        cells=sum(len(row) for row in COMPLETE_GRID),
                    )
                    if any(
                        selected_page._text_selection is not None
                        or selected_page._text_selection_source_ranges
                        for selected_page in window.document._pages
                    ):
                        return finish(
                            window,
                            "the whole-table selection remained visible after its "
                            "clipboard copy succeeded",
                        )
                    if (
                        not window._notification_revealer.get_reveal_child()
                        or window._notification_title.get_text() != "Table copied"
                        or window._notification_detail.get_text()
                        != "The complete table is ready to paste"
                    ):
                        return finish(
                            window,
                            "copying a complete table did not show the standard "
                            "success notification",
                        )
                    if arguments.case == "table-copy":
                        print(
                            "hovering the table reveals a control that copies its "
                            "complete structured contents"
                        )
                        return finish(window)
                    state["phase"] = "single_cell"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "single_cell":
                    # A drag wholly inside one wrapped cell is a text-range
                    # selection in a single table cell, never a row-major
                    # selection through neighbouring cells.
                    _drag_characters(
                        page, client, desktop_input, "selected", "words"
                    )
                    state["phase"] = "single_cell_copy"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "single_cell_copy":
                    _clipboard_odt(text="selected words", rows=1, cells=1)
                    if arguments.case == "single-cell":
                        print("an intra-cell drag copies only its selected text")
                        return finish(window)
                    state["phase"] = "double_click_first"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "double_click_first":
                    _click(page, client, desktop_input, "C3 outer-right")
                    state["second_click_at"] = time.monotonic() + 0.12
                    state["phase"] = "double_click_second"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "double_click_second":
                    if time.monotonic() < float(state["second_click_at"]):
                        return GLib.SOURCE_CONTINUE
                    _click(page, client, desktop_input, "C3 outer-right")
                    state["double_click_check_at"] = time.monotonic() + 0.25
                    state["phase"] = "double_click_copy"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "double_click_copy":
                    if time.monotonic() < float(state["double_click_check_at"]):
                        return GLib.SOURCE_CONTINUE
                    _clipboard_odt(text=GRID[2][2], rows=1, cells=1)
                    if arguments.case == "double-click":
                        print("double-clicking a table cell copies its complete contents")
                        return finish(window)
                    state["phase"] = "column"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "column":
                    # Match the reported upward gesture.  Despite the PDF's
                    # row-major glyph stream, only the centre column belongs
                    # to this rectangular table selection.
                    _drag_characters(page, client, desktop_input, "B3", "B1")
                    state["phase"] = "column_shape"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "column_shape":
                    _clipboard_odt(
                        text="\n".join(row[1] for row in GRID), rows=3, cells=3
                    )
                    if arguments.case == "column":
                        print("an upward drag copies only its table column")
                        return finish(window)
                    state["phase"] = "row"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "row":
                    _drag_characters(page, client, desktop_input, "A2", "C2")
                    state["phase"] = "row_copy"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "row_copy":
                    _clipboard_odt(text="\t".join(GRID[1]), rows=1, cells=3)
                    start_probe("row")
                    state["phase"] = "row_paste"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "row_paste":
                    complete, error = completed_probe([GRID[1]])
                    if not complete:
                        return GLib.SOURCE_CONTINUE
                    if error:
                        return finish(window, error)
                    _drag_characters(page, client, desktop_input, "B3", "B1")
                    state["phase"] = "column_copy"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "column_copy":
                    _clipboard_odt(
                        text="\n".join(row[1] for row in GRID), rows=3, cells=3
                    )
                    start_probe("column")
                    state["phase"] = "column_paste"
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "column_paste":
                    complete, error = completed_probe([[row[1]] for row in GRID])
                    if not complete:
                        return GLib.SOURCE_CONTINUE
                    if error:
                        return finish(window, error)
                    print(
                        "table drags copy exact cell grids whose live clipboard "
                        "creates coherent Writer tables"
                    )
                    return finish(window)
            except (OSError, RuntimeError, subprocess.SubprocessError, zipfile.BadZipFile) as error:
                return finish(window, str(error))
            return GLib.SOURCE_CONTINUE

        def activate(app: Gtk.Application) -> None:
            window = DocxWindow(app, fixture)
            window.set_default_size(1440, 900)
            window.show_all()
            GLib.timeout_add(100, check, window)
            GLib.timeout_add(
                TIMEOUT_MS,
                finish,
                window,
                "timed out verifying table selection and LibreOffice paste",
            )

        application.connect("activate", activate)
        application.run(["see-docx-rich-selection-smoke"])
    if error := result.get("error"):
        raise RuntimeError(error)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
