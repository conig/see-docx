#!/usr/bin/env python3
"""Exercise a real watched DOCX refresh and assert the reading location holds.

This is deliberately a manual smoke test: it needs a GTK display, Poppler, and
LibreOffice. It copies the supplied document to a private temporary file,
touches only that copy, and exits non-zero if the refresh returns to the top.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import tempfile

import gi

gi.require_version("Gtk", "3.0")

from gi.repository import Gio, GLib, Gtk

from see_docx.viewer import APPLICATION_ID, DocxWindow


TIMEOUT_MS = 20_000


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path, help="A multi-page DOCX fixture")
    arguments = parser.parse_args()
    source = arguments.path.resolve()
    if not source.is_file():
        parser.error(f"No such DOCX: {source}")

    with tempfile.TemporaryDirectory(prefix="see-docx-ui-smoke-") as directory:
        watched = Path(directory) / "watched.docx"
        shutil.copy2(source, watched)
        result: dict[str, str] = {}
        state: dict[str, object] = {
            "replaced": False,
            "expected": None,
            "initial_revision": None,
            "before": None,
        }
        application = Gtk.Application(
            application_id=f"{APPLICATION_ID}.smoke",
            flags=Gio.ApplicationFlags.NON_UNIQUE,
        )

        def finish(window: DocxWindow, message: str | None = None) -> bool:
            if message:
                result["error"] = message
            window.close()
            application.quit()
            return GLib.SOURCE_REMOVE

        def check(window: DocxWindow) -> bool:
            if not window.document.has_document:
                return GLib.SOURCE_CONTINUE
            adjustment = window.document.widget.get_vadjustment()
            maximum = adjustment.get_upper() - adjustment.get_lower() - adjustment.get_page_size()
            if maximum <= 40:
                # The document object exists before GTK has finished measuring
                # its replacement page stack; wait for a real scroll range.
                return GLib.SOURCE_CONTINUE

            if not state["replaced"]:
                adjustment.set_value(adjustment.get_lower() + maximum * 0.62)
                state["expected"] = window.document.capture_position()
                state["initial_revision"] = window._rendered_revision
                state["before"] = (
                    adjustment.get_value(),
                    maximum,
                    window.document._page_geometries(),
                )
                replacement = watched.with_name("replacement.docx")
                shutil.copy2(watched, replacement)
                replacement.replace(watched)
                state["replaced"] = True
                return GLib.SOURCE_CONTINUE

            if window._rendered_revision == state["initial_revision"]:
                return GLib.SOURCE_CONTINUE
            if (
                window._process is not None
                or window.document.restore_pending
                or not window.status.get_text().startswith("Live preview updated")
            ):
                return GLib.SOURCE_CONTINUE

            expected = state["expected"]
            actual = window.document.capture_position()
            if expected is None or expected.page_index != actual.page_index:
                return finish(
                    window,
                    f"page changed from {expected} to {actual}; before={state['before']}",
                )
            if abs((expected.page_fraction or 0.0) - (actual.page_fraction or 0.0)) > 0.03:
                return finish(
                    window,
                    "within-page offset changed from " f"{expected} to {actual}; before={state['before']}",
                )
            print(f"position preserved: {actual}")
            return finish(window)

        def activate(_application: Gtk.Application) -> None:
            window = DocxWindow(application, watched)
            window.show_all()
            GLib.timeout_add(100, check, window)
            GLib.timeout_add(
                TIMEOUT_MS,
                lambda: finish(
                    window,
                    "timed out waiting for the live refresh "
                    f"(revision={window._revision}; status={window.status.get_text()!r})",
                ),
            )

        application.connect("activate", activate)
        application.run(["see-docx-ui-smoke"])
        if error := result.get("error"):
            print(f"FAIL: {error}")
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
