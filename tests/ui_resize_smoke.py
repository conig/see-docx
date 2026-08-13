#!/usr/bin/env python3
"""Resize a realized viewer and verify its page/canvas proportion stays stable."""

from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import tempfile

import gi

gi.require_version("Gdk", "3.0")
gi.require_version("Gtk", "3.0")

from gi.repository import Gdk, Gio, GLib, Gtk

from see_docx.viewer import APPLICATION_ID, DocxWindow


SMALL_SIZE = (700, 600)
TIMEOUT_MS = 25_000
SMOKE_TEST_CLASS = "codex-smoke-test"


def main() -> int:
    GLib.set_prgname(SMOKE_TEST_CLASS)
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path, help="A DOCX fixture")
    arguments = parser.parse_args()
    source = arguments.path.resolve()
    if not source.is_file():
        parser.error(f"No such DOCX: {source}")

    with tempfile.TemporaryDirectory(prefix="see-docx-resize-smoke-") as directory:
        watched = Path(directory) / "watched.docx"
        shutil.copy2(source, watched)
        result: dict[str, str] = {}
        state: dict[str, object] = {"stage": "small"}
        application = Gtk.Application(
            application_id=f"{APPLICATION_ID}.resize-smoke",
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
            width = window.get_allocated_width()
            if width <= 0:
                return GLib.SOURCE_CONTINUE
            page = window.document._pages[0]
            ratio = page._zoom / width
            stage = state["stage"]

            if stage == "small":
                state["small_width"] = width
                state["small_zoom"] = page._zoom
                state["small_ratio"] = ratio
                state["stage"] = "large"
                window.fullscreen()
                return GLib.SOURCE_CONTINUE

            if stage == "large":
                small_width = int(state["small_width"])
                if width <= small_width * 1.25:
                    return GLib.SOURCE_CONTINUE
                expected = float(state["small_ratio"])
                if abs(ratio - expected) > expected * 0.03:
                    return finish(
                        window,
                        "growing the window changed the page/canvas width proportion "
                        f"(small={small_width}px@{float(state['small_zoom']):.3f}, "
                        f"large={width}px@{page._zoom:.3f})",
                    )
                state["large_width"] = width
                state["stage"] = "small-again"
                window.unfullscreen()
                return GLib.SOURCE_CONTINUE

            large_width = int(state["large_width"])
            if width >= large_width * 0.80:
                return GLib.SOURCE_CONTINUE
            expected = float(state["small_ratio"])
            if abs(ratio - expected) > expected * 0.03:
                return finish(
                    window,
                    "shrinking the window changed the page/canvas width proportion",
                )
            if abs(page._zoom - float(state["small_zoom"])) > 0.02:
                return finish(window, "shrinking did not restore the original zoom")
            print("window growth and shrinkage preserve page/canvas proportions")
            return finish(window)

        def activate(_application: Gtk.Application) -> None:
            window = DocxWindow(application, watched)
            window.set_type_hint(Gdk.WindowTypeHint.DIALOG)
            window.set_default_size(*SMALL_SIZE)
            window.show_all()
            GLib.timeout_add(100, check, window)
            GLib.timeout_add(
                TIMEOUT_MS,
                lambda: finish(window, "timed out waiting for both resize events"),
            )

        application.connect("activate", activate)
        application.run(["see-docx-resize-smoke"])
        if error := result.get("error"):
            print(f"FAIL: {error}")
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
