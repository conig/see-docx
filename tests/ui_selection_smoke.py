#!/usr/bin/env python3
"""Exercise an actual GTK pointer-selection drag across a PDF page break."""

from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import tempfile
from types import SimpleNamespace

import gi

gi.require_version("Gdk", "3.0")
gi.require_version("Gtk", "3.0")

from gi.repository import Gdk, Gio, GLib, Gtk

from see_docx.viewer import APPLICATION_ID, DocxWindow


TIMEOUT_MS = 20_000


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path, help="A multi-page DOCX fixture")
    arguments = parser.parse_args()
    source = arguments.path.resolve()
    if not source.is_file():
        parser.error(f"No such DOCX: {source}")

    with tempfile.TemporaryDirectory(prefix="see-docx-selection-smoke-") as directory:
        watched = Path(directory) / "watched.docx"
        shutil.copy2(source, watched)
        result: dict[str, str] = {}
        state: dict[str, object] = {"edge_drag_started": False, "attempts": 0}
        application = Gtk.Application(
            application_id=f"{APPLICATION_ID}.selection-smoke",
            flags=Gio.ApplicationFlags.NON_UNIQUE,
        )

        def finish(window: DocxWindow, message: str | None = None) -> bool:
            if message:
                result["error"] = message
            window.close()
            application.quit()
            return GLib.SOURCE_REMOVE

        def check(window: DocxWindow) -> bool:
            if not window.document.has_document or window.document.page_count < 2:
                return GLib.SOURCE_CONTINUE
            adjustment = window.document.widget.get_vadjustment()
            if window.document._maximum_scroll() <= 80:
                return GLib.SOURCE_CONTINUE

            first_page, second_page = window.document._pages[:2]
            if not state["edge_drag_started"]:
                start = SimpleNamespace(button=1, x=120.0, y=180.0)
                beyond_first_page = SimpleNamespace(
                    x=310.0,
                    y=first_page.get_allocated_height() + 120.0,
                )
                state["scroll_before_edge_drag"] = adjustment.get_value()
                first_page._on_button_press(first_page, start)
                first_page._on_motion(first_page, beyond_first_page)
                if second_page._text_selection is None:
                    return finish(
                        window,
                        "dragging beyond a PDF page must continue the visible selection on the next page",
                    )
                state["edge_drag_started"] = True
                return GLib.SOURCE_CONTINUE

            if adjustment.get_value() <= state["scroll_before_edge_drag"]:
                state["attempts"] += 1
                if state["attempts"] < 10:
                    return GLib.SOURCE_CONTINUE
                return finish(
                    window,
                    "holding a drag beyond a PDF page must automatically scroll the viewport",
                )

            start = SimpleNamespace(button=1, x=120.0, y=180.0)
            first_page._on_button_release(
                first_page,
                SimpleNamespace(button=1, x=310.0, y=240.0),
            )
            first_page._on_button_press(first_page, start)
            first_page._on_motion(
                first_page, SimpleNamespace(x=310.0, y=240.0)
            )
            before = first_page._text_selection
            scroll_before = adjustment.get_value()
            if not first_page._on_scroll(
                first_page,
                SimpleNamespace(direction=Gdk.ScrollDirection.DOWN, state=0),
            ) or adjustment.get_value() <= scroll_before:
                return finish(
                    window,
                    "a wheel event over the rendered page must scroll an active selection",
                )
            after = first_page._text_selection
            if before is None or after is None or after.bottom <= before.bottom:
                return finish(
                    window,
                    "a wheel scroll must extend the held text selection",
                )
            print("cross-page selection and scrolling drag work")
            return finish(window)

        def activate(_application: Gtk.Application) -> None:
            window = DocxWindow(application, watched)
            window.show_all()
            GLib.timeout_add(100, check, window)
            GLib.timeout_add(
                TIMEOUT_MS,
                lambda: finish(window, "timed out waiting for a rendered multi-page preview"),
            )

        application.connect("activate", activate)
        application.run(["see-docx-selection-smoke"])
        if error := result.get("error"):
            print(f"FAIL: {error}")
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
