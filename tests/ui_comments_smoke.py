#!/usr/bin/env python3
"""Verify that focused comments pop out beside the realized document page."""

from __future__ import annotations

from pathlib import Path
import subprocess
import tempfile
from types import SimpleNamespace
import zipfile

import gi

gi.require_version("Gdk", "3.0")
gi.require_version("Gtk", "3.0")

from gi.repository import Gdk, Gio, GLib, Gtk

from see_docx.viewer import APPLICATION_ID, DocxWindow


TIMEOUT_MS = 90_000
SMOKE_TEST_CLASS = "codex-smoke-test"
TARGET_TEXT = b"First section"
COMMENT_TEXT = (
    "Keep the active comment close to the page so it can be read in context."
)


def _commented_fixture(destination: Path) -> None:
    """Build a valid DOCX and attach one Word comment to visible body text."""

    source = Path(__file__).with_name("fixtures") / "live_refresh.md"
    subprocess.run(
        ["pandoc", str(source), "-o", str(destination)],
        check=True,
        capture_output=True,
        text=True,
    )
    with zipfile.ZipFile(destination) as archive:
        members = {name: archive.read(name) for name in archive.namelist()}

    document_xml = members["word/document.xml"]
    run = b'<w:r><w:t xml:space="preserve">' + TARGET_TEXT + b"</w:t></w:r>"
    if document_xml.count(run) != 1:
        raise RuntimeError("the smoke fixture no longer has one comment target")
    members["word/document.xml"] = document_xml.replace(
        run,
        b'<w:commentRangeStart w:id="7"/>'
        + run
        + b'<w:commentRangeEnd w:id="7"/>'
        + b'<w:r><w:rPr><w:commentReference w:id="7"/></w:rPr></w:r>',
        1,
    )
    comments_xml = members["word/comments.xml"]
    if not comments_xml.rstrip().endswith(b" />"):
        raise RuntimeError("pandoc comments.xml changed shape")
    members["word/comments.xml"] = comments_xml.rstrip()[:-3] + (
        b'><w:comment w:id="7" w:author="Reader" w:initials="R">'
        b"<w:p><w:r><w:t>"
        + COMMENT_TEXT.encode()
        + b"</w:t></w:r></w:p></w:comment></w:comments>"
    )

    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)


def main() -> int:
    # Hyprland routes this stable initial app ID to workspace 15 before map.
    GLib.set_prgname(SMOKE_TEST_CLASS)
    with tempfile.TemporaryDirectory(
        prefix="see-docx-comments-smoke-"
    ) as directory:
        fixture = Path(directory) / "commented.docx"
        _commented_fixture(fixture)
        result: dict[str, str] = {}
        state: dict[str, object] = {"phase": "focus", "checks": 0}
        application = Gtk.Application(
            application_id=f"{APPLICATION_ID}.comments-smoke",
            flags=Gio.ApplicationFlags.NON_UNIQUE,
        )

        def finish(window: DocxWindow, message: str | None = None) -> bool:
            if message:
                result["error"] = message
            window.close()
            application.quit()
            return GLib.SOURCE_REMOVE

        def check(window: DocxWindow) -> bool:
            if not window.document.has_document or not window._comment_annotations:
                return GLib.SOURCE_CONTINUE
            if state["phase"] == "focus":
                state["phase"] = "popout"
                if not window._on_key_press(
                    window,
                    SimpleNamespace(keyval=Gdk.KEY_c, state=0),
                ):
                    return finish(window, "c did not focus the comment rail")
                return GLib.SOURCE_CONTINUE

            state["checks"] = int(state["checks"]) + 1
            rail_card = window._comment_cards[window._active_comment_id]
            rail_style = rail_card.get_style_context()
            popout = window._active_comment_float_card
            layer_parent = window._active_comment_layer.get_parent()
            if state["phase"] == "restored":
                layer_allocation = window._active_comment_layer.get_allocation()
                if (
                    window._active_comment_float_card is None
                    and not rail_style.has_class("comment-rail-ghost")
                    and window._active_comment_layer.get_visible()
                    and isinstance(layer_parent, Gtk.Overlay)
                    and layer_parent.get_overlay_pass_through(
                        window._active_comment_layer
                    )
                    and layer_allocation.width > 1
                    and layer_allocation.height > 1
                ):
                    print(
                        "focused comment is visible beside the page; "
                        "the inactive layer remains allocated and passes input through"
                    )
                    return finish(window)
                if int(state["checks"]) < 30:
                    return GLib.SOURCE_CONTINUE
                return finish(
                    window,
                    "closing comment focus did not restore a measurable pass-through layer",
                )
            if (
                popout is not None
                and popout.get_visible()
                and popout.get_mapped()
                and window._active_comment_layer.get_visible()
                and rail_style.has_class("comment-rail-ghost")
            ):
                # The promoted card must be interactive without turning its
                # full-workspace positioning layer into an input shield. That
                # shield suppresses the PDF's text cursor, wheel scrolling,
                # and pointer-selection events everywhere outside the card.
                if not isinstance(layer_parent, Gtk.Overlay):
                    return finish(window, "comment layer is not in an overlay")
                if not layer_parent.get_overlay_pass_through(
                    window._active_comment_layer
                ):
                    return finish(
                        window,
                        "focused comment blocks document cursor, scrolling, and selection",
                    )
                popout_parent = popout.get_parent()
                if not isinstance(popout_parent, Gtk.Overlay):
                    return finish(
                        window,
                        "focused comment is not an independently interactive overlay",
                    )
                if popout_parent.get_overlay_pass_through(popout):
                    return finish(window, "focused comment does not accept pointer input")
                rail_allocation = rail_card.get_allocation()
                popout_allocation = popout.get_allocation()
                if popout_allocation.width <= 0 or popout_allocation.height <= 0:
                    return finish(
                        window,
                        "the page-adjacent comment has no visible allocation",
                    )
                if rail_allocation.width <= 0 or rail_allocation.height <= 0:
                    return finish(
                        window,
                        "the rail ghost no longer reserves the card's space",
                    )
                if not window._on_key_press(
                    window,
                    SimpleNamespace(keyval=Gdk.KEY_Escape, state=0),
                ):
                    return finish(window, "Esc did not return focus to the document")
                state["phase"] = "restored"
                state["checks"] = 0
                return GLib.SOURCE_CONTINUE
            if int(state["checks"]) < 30:
                return GLib.SOURCE_CONTINUE
            context = window._comment_float_context(window._active_comment_id)
            rail_allocation = rail_card.get_allocation()
            layer_allocation = window._active_comment_layer.get_allocation()
            geometry = window._comment_float_geometry_for_thread(
                window._active_comment_id,
                rail_allocation.height,
            )
            return finish(
                window,
                "focused comment did not keep a rail ghost and a mapped page-adjacent copy: "
                f"focused={window._comments_focused}; "
                f"float_id={window._active_comment_float_id!r}; "
                f"popout={popout!r}; "
                f"layer_visible={window._active_comment_layer.get_visible()}; "
                f"ghost={rail_style.has_class('comment-rail-ghost')}; "
                f"window={window.get_allocated_width()}x{window.get_allocated_height()}; "
                f"layer={layer_allocation.width}x{layer_allocation.height}; "
                f"rail_card={rail_allocation.width}x{rail_allocation.height}; "
                f"context={context!r}; "
                f"geometry={geometry!r}",
            )

        def activate(app: Gtk.Application) -> None:
            window = DocxWindow(app, fixture)
            window.set_default_size(1440, 900)
            window.show_all()
            GLib.timeout_add(100, check, window)
            GLib.timeout_add(
                TIMEOUT_MS,
                finish,
                window,
                "timed out waiting for the focused comment popout",
            )

        application.connect("activate", activate)
        application.run(["see-docx-comments-smoke"])

    if "error" in result:
        raise RuntimeError(result["error"])
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
