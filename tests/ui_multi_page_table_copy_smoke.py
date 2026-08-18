#!/usr/bin/env python3
"""Copy one complete table from its hover control on a later page segment."""

from __future__ import annotations

from pathlib import Path
import subprocess
import tempfile
import time

import gi

gi.require_version("Gio", "2.0")
gi.require_version("Gtk", "3.0")

from gi.repository import Gio, GLib, Gtk

from see_docx.viewer import APPLICATION_ID, DocxWindow
from sway_test_support import (
    DesktopInput,
    WORKSPACE,
    capture_failure_screenshot,
    focus_client,
    focus_workspace,
    smoke_client,
)
from ui_rich_selection_smoke import _clipboard_odt


TIMEOUT_MS = 120_000
SMOKE_TEST_CLASS = "codex-smoke-test"
TABLE = [
    ["Record", "Delivery pathway", "Verification evidence"],
    *[
        [
            f"Item {index:02d}",
            f"Programme item {index:02d} uses routine review meetings, staff "
            "feedback and locally agreed actions to respond to changing needs.",
            f"Evidence item {index:02d} includes activity records, response times, "
            "implementation notes, costs and independently reviewed outcomes.",
        ]
        for index in range(1, 35)
    ],
]


def _fixture(destination: Path) -> None:
    markdown = destination.with_suffix(".md")
    markdown.write_text(
        "\n".join(
            (
                "| " + " | ".join(TABLE[0]) + " |",
                "|" + "|".join(["---"] * len(TABLE[0])) + "|",
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


def main() -> int:
    GLib.set_prgname(SMOKE_TEST_CLASS)
    with tempfile.TemporaryDirectory(
        prefix="see-docx-multi-page-table-copy-smoke-"
    ) as directory:
        fixture = Path(directory) / "multi-page-table.docx"
        _fixture(fixture)
        result: dict[str, str] = {}
        state: dict[str, object] = {"phase": "setup"}
        desktop_input: DesktopInput | None = None
        application = Gtk.Application(
            application_id=f"{APPLICATION_ID}.multi-page-table-copy-smoke",
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
                source = window.document._rich_source
                client = smoke_client(SMOKE_TEST_CLASS)
                if source is None or client is None:
                    return GLib.SOURCE_CONTINUE
                table_index = next(
                    fragment.table_index
                    for fragment in source._fragments
                    if fragment.table_index is not None
                )
                segments = [
                    (
                        page_index,
                        page,
                        [
                            layout
                            for layout in page._table_cell_layouts
                            if layout.fragment.table_index == table_index
                        ],
                    )
                    for page_index, page in enumerate(window.document._pages)
                ]
                segments = [segment for segment in segments if segment[2]]
                if len(segments) < 2:
                    return GLib.SOURCE_CONTINUE
                _page_index, page, layouts = segments[1]

                if desktop_input is None:
                    if (
                        window.document._maximum_scroll() <= 0
                        or page.get_allocation().y
                        <= window.document._pages[0].get_allocation().y
                    ):
                        return GLib.SOURCE_CONTINUE
                    desktop_input = DesktopInput(directory)
                    focus_workspace(WORKSPACE)
                    focus_client(client)
                    adjustment = window.document.widget.get_vadjustment()
                    page_top = page.get_allocation().y
                    segment_top = min(layout.top for layout in layouts) * page._zoom
                    target = (
                        page_top
                        + segment_top
                        - adjustment.get_page_size() * 0.28
                    )
                    adjustment.set_value(
                        min(
                            max(target, adjustment.get_lower()),
                            adjustment.get_lower()
                            + window.document._maximum_scroll(),
                        )
                    )
                    state["phase"] = "scroll"
                    state["ready_at"] = time.monotonic() + 0.8
                    state["hover_attempts"] = 0
                    return GLib.SOURCE_CONTINUE
                if time.monotonic() < float(state["ready_at"]):
                    return GLib.SOURCE_CONTINUE

                window_x, window_y = client["at"]
                translated = page.translate_coordinates(window, 0, 0)
                if not translated[0]:
                    return GLib.SOURCE_CONTINUE
                page_x, page_y = translated[-2:]
                if state["phase"] == "scroll":
                    if (
                        page_y + page.get_allocated_height() <= 0
                        or page_y >= window.get_allocated_height()
                    ):
                        adjustment = window.document.widget.get_vadjustment()
                        target = (
                            page.get_allocation().y
                            + min(layout.top for layout in layouts) * page._zoom
                            - adjustment.get_page_size() * 0.28
                        )
                        adjustment.set_value(
                            min(
                                max(target, adjustment.get_lower()),
                                adjustment.get_lower()
                                + window.document._maximum_scroll(),
                            )
                        )
                        state["ready_at"] = time.monotonic() + 0.2
                        return GLib.SOURCE_CONTINUE
                    state["phase"] = "hover"
                    return GLib.SOURCE_CONTINUE
                if state["phase"] == "hover":
                    target_layout = next(
                        (
                            layout
                            for layout in layouts
                            if (layout.fragment.row_index or 0) > 0
                        ),
                        layouts[0],
                    )
                    target_x = (
                        (target_layout.left + target_layout.right)
                        * page._zoom
                        / 2
                    )
                    target_y = (
                        (target_layout.top + target_layout.bottom)
                        * page._zoom
                        / 2
                    )
                    state["hover_origin"] = (target_x, target_y)
                    attempts = int(state["hover_attempts"])
                    desktop_input.move_cursor(
                        int(window_x + page_x + target_x + attempts % 2),
                        int(window_y + page_y + target_y),
                    )
                    state["hover_attempts"] = attempts + 1
                    state["phase"] = "button"
                    state["ready_at"] = time.monotonic() + 0.15
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "button":
                    hovered = page._table_copy_hover_index
                    if hovered is None:
                        if int(state["hover_attempts"]) < 5:
                            state["phase"] = "hover"
                            return GLib.SOURCE_CONTINUE
                        return finish(
                            window,
                            "hovering a later table page did not reveal its copy control",
                        )
                    button = page._table_copy_button(hovered)
                    if button is None:
                        return finish(
                            window, "the later-page table control has no hit box"
                        )
                    table_left = min(layout.left for layout in layouts) * page._zoom
                    table_top = min(layout.top for layout in layouts) * page._zoom
                    if (
                        button.right > table_left - 4.0
                        or button.bottom > table_top - 4.0
                    ):
                        return finish(
                            window,
                            "the later-page copy control is not wholly outside "
                            "the table segment",
                        )
                    target_x = (button.left + button.right) / 2
                    target_y = (button.top + button.bottom) / 2
                    origin_x, origin_y = state["hover_origin"]
                    state["transition"] = [
                        (
                            origin_x + (target_x - origin_x) * step / 32,
                            origin_y + (target_y - origin_y) * step / 32,
                        )
                        for step in range(1, 33)
                    ]
                    state["transition_step"] = 0
                    state["phase"] = "transition"
                    state["ready_at"] = time.monotonic() + 0.04
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "transition":
                    step = int(state["transition_step"])
                    if step and page._table_copy_hover_index is None:
                        return finish(
                            window,
                            "the later-page copy control disappeared while the "
                            "pointer travelled to it",
                        )
                    points = state["transition"]
                    if step < len(points):
                        target_x, target_y = points[step]
                        desktop_input.move_cursor(
                            int(window_x + page_x + target_x),
                            int(window_y + page_y + target_y),
                        )
                        state["transition_step"] = step + 1
                        state["ready_at"] = time.monotonic() + 0.04
                        return GLib.SOURCE_CONTINUE
                    state["phase"] = "click"
                    state["ready_at"] = time.monotonic() + 0.15
                    return GLib.SOURCE_CONTINUE

                if state["phase"] == "click":
                    if not page._table_copy_button_hot:
                        return finish(
                            window,
                            "the later-page table control disappeared before clicking",
                        )
                    desktop_input.left_button(pressed=True)
                    desktop_input.left_button(pressed=False)
                    state["phase"] = "check"
                    state["ready_at"] = time.monotonic() + 0.3
                    return GLib.SOURCE_CONTINUE

                expected_cells = sum(len(row) for row in TABLE)
                _clipboard_odt(
                    text="\n".join("\t".join(row) for row in TABLE),
                    rows=len(TABLE),
                    cells=expected_cells,
                )
                selected_segments = [
                    index
                    for index, segment_page, _segment_layouts in segments
                    if segment_page._text_selection is not None
                    or segment_page._text_selection_source_ranges
                ]
                if selected_segments:
                    return finish(
                        window,
                        "copying from a later segment left the multi-page table "
                        f"selected on pages {selected_segments!r}",
                    )
                print(
                    "a hover control on a later page copies one complete multi-page "
                    "table and clears its selection"
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
                "timed out verifying multi-page table hover copy",
            )

        application.connect("activate", activate)
        application.run(["see-docx-multi-page-table-copy-smoke"])
    if error := result.get("error"):
        raise RuntimeError(error)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
