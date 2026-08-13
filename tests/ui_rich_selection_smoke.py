#!/usr/bin/env python3
"""Verify a real PDF drag exports a formatted DOCX table cell as HTML."""

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
TARGET_TEXT = "Passed"
SMOKE_TEST_CLASS = "codex-smoke-test"


def main() -> int:
    GLib.set_prgname(SMOKE_TEST_CLASS)
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path, help="A DOCX fixture containing bold 'Passed'")
    arguments = parser.parse_args()
    source = arguments.path.resolve()
    if not source.is_file():
        parser.error(f"No such DOCX: {source}")

    with tempfile.TemporaryDirectory(prefix="see-docx-rich-selection-smoke-") as directory:
        watched = Path(directory) / "watched.docx"
        shutil.copy2(source, watched)
        result: dict[str, str] = {}
        application = Gtk.Application(
            application_id=f"{APPLICATION_ID}.rich-selection-smoke",
            flags=Gio.ApplicationFlags.NON_UNIQUE,
        )

        def finish(window: DocxWindow, message: str | None = None) -> bool:
            if message:
                result["error"] = message
            window.close()
            application.quit()
            return GLib.SOURCE_REMOVE

        def check(window: DocxWindow) -> bool:
            if not window.document.has_document or not window.document._pages:
                return GLib.SOURCE_CONTINUE
            page = window.document._pages[0]
            text = page._page.get_text()
            target_index = text.find(TARGET_TEXT)
            if target_index < 0:
                return finish(window, f"rendered PDF does not contain {TARGET_TEXT!r}")
            has_layout, rectangles = page._page.get_text_layout()
            if not has_layout or target_index + len(TARGET_TEXT) > len(rectangles):
                return finish(window, "rendered PDF has no glyph layout for the target cell")
            first = rectangles[target_index]
            last = rectangles[target_index + len(TARGET_TEXT) - 1]
            start = SimpleNamespace(
                button=1,
                x=(float(first.x1) + 0.5) * page._zoom,
                y=(float(first.y1) + float(first.y2)) * page._zoom / 2,
            )
            end = SimpleNamespace(
                button=1,
                x=(float(last.x2) - 0.5) * page._zoom,
                y=(float(last.y1) + float(last.y2)) * page._zoom / 2,
            )
            page._on_button_press(page, start)
            page._on_motion(page, end)
            page._on_button_release(page, end)

            clipboard = Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD)
            html_data = clipboard.wait_for_contents(
                Gdk.Atom.intern("text/html", False)
            )
            if html_data is None:
                return finish(window, "selection did not offer text/html to the clipboard")
            html = bytes(html_data.get_data()).decode("utf-8")
            if "<table" not in html or "<strong>Passed</strong>" not in html:
                return finish(
                    window,
                    "the copied table cell lost its table boundary or bold formatting",
                )
            if "Alpha" in html:
                return finish(window, "copying one cell included its neighbouring cell")
            if clipboard.wait_for_text() != TARGET_TEXT:
                return finish(window, "the plain-text clipboard fallback differs from the cell")
            print("rich table-cell selection offers matching HTML and text clipboard data")
            return finish(window)

        def activate(_application: Gtk.Application) -> None:
            window = DocxWindow(application, watched)
            window.show_all()
            GLib.timeout_add(100, check, window)
            GLib.timeout_add(
                TIMEOUT_MS,
                lambda: finish(window, "timed out waiting for a rendered preview"),
            )

        application.connect("activate", activate)
        application.run(["see-docx-rich-selection-smoke"])
        if error := result.get("error"):
            print(f"FAIL: {error}")
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
