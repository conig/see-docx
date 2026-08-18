#!/usr/bin/env python3
"""Verify that focused comments pop out beside the realized document page."""

from __future__ import annotations

from pathlib import Path
import subprocess
import tempfile
import zipfile

import gi

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
    sway_command,
)


TIMEOUT_MS = 90_000
SMOKE_TEST_CLASS = "codex-smoke-test"
SHORT_COMMENT_TARGET = b"First section"
LONG_COMMENT_TARGET = b"Repeated material"
SHORT_COMMENT_TEXT = "Clarify this."
LONG_COMMENT_TEXT = (
    "Keep the active comment close to the page so it can be read in context. "
    "This deliberately longer note checks that wrapped prose uses the same "
    "stable width as a short note instead of distorting the document page. "
    "The width must stay constant as keyboard focus moves between comments."
)
INPUT_SETTLE_CHECKS = 15


def _page_drag_points(page: object) -> tuple[tuple[float, float], tuple[float, float]]:
    """Choose two visible glyph centres so real pointer input selects text."""

    text = page._page.get_text()
    has_layout, rectangles = page._page.get_text_layout()
    if not has_layout or len(text) != len(rectangles):
        raise RuntimeError("the rendered comment page has no aligned glyph layout")
    candidates = [
        index
        for index, character in enumerate(text)
        if not character.isspace()
        and page._selection_flow_map.get(index, "main") == "main"
    ]
    for position, start_index in enumerate(candidates):
        start = rectangles[start_index]
        start_y = (float(start.y1) + float(start.y2)) / 2
        same_line = [
            index
            for index in candidates[position + 8 :]
            if abs(
                (float(rectangles[index].y1) + float(rectangles[index].y2)) / 2
                - start_y
            )
            < 2.0
        ]
        if same_line:
            end = rectangles[same_line[min(12, len(same_line) - 1)]]
            return (
                (
                    (float(start.x1) + float(start.x2)) * page._zoom / 2,
                    start_y * page._zoom,
                ),
                (
                    (float(end.x1) + float(end.x2)) * page._zoom / 2,
                    (float(end.y1) + float(end.y2)) * page._zoom / 2,
                ),
            )
    raise RuntimeError("the rendered comment page has no selectable text line")


def _commented_fixture(destination: Path) -> None:
    """Build a valid DOCX with short and long comments on visible text."""

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
    targets = ((b"7", SHORT_COMMENT_TARGET), (b"8", LONG_COMMENT_TARGET))
    for comment_id, target in targets:
        run = b'<w:r><w:t xml:space="preserve">' + target + b"</w:t></w:r>"
        if document_xml.count(run) != 1:
            raise RuntimeError("the smoke fixture no longer has one comment target")
        document_xml = document_xml.replace(
            run,
            b'<w:commentRangeStart w:id="'
            + comment_id
            + b'"/>'
            + run
            + b'<w:commentRangeEnd w:id="'
            + comment_id
            + b'"/>'
            + b'<w:r><w:rPr><w:commentReference w:id="'
            + comment_id
            + b'"/></w:rPr></w:r>',
            1,
        )
    members["word/document.xml"] = document_xml
    comments_xml = members["word/comments.xml"]
    if not comments_xml.rstrip().endswith(b" />"):
        raise RuntimeError("pandoc comments.xml changed shape")
    members["word/comments.xml"] = comments_xml.rstrip()[:-3] + (
        b'><w:comment w:id="7" w:author="Reader" w:initials="R">'
        b"<w:p><w:r><w:t>"
        + SHORT_COMMENT_TEXT.encode()
        + b"</w:t></w:r></w:p></w:comment>"
        b'<w:comment w:id="8" w:author="Reader" w:initials="R">'
        b"<w:p><w:r><w:t>"
        + LONG_COMMENT_TEXT.encode()
        + b"</w:t></w:r></w:p></w:comment></w:comments>"
    )

    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)


def main() -> int:
    # The headless session starts on workspace 15 before this surface maps.
    GLib.set_prgname(SMOKE_TEST_CLASS)
    with tempfile.TemporaryDirectory(
        prefix="see-docx-comments-smoke-"
    ) as directory:
        fixture = Path(directory) / "commented.docx"
        _commented_fixture(fixture)
        result: dict[str, str] = {}
        state: dict[str, object] = {"phase": "focus", "checks": 0}
        desktop_input: DesktopInput | None = None
        application = Gtk.Application(
            application_id=f"{APPLICATION_ID}.comments-smoke",
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
            if not window.document.has_document:
                return GLib.SOURCE_CONTINUE
            if len(window._comment_annotations) != 2:
                return finish(window, "the smoke fixture did not load both comments")
            if state["phase"] == "focus":
                try:
                    desktop_input = DesktopInput(directory)
                    client = smoke_client(SMOKE_TEST_CLASS)
                    if client is None:
                        return finish(window, "Sway did not expose the smoke-test window")
                    if client["workspace"]["id"] != int(WORKSPACE):
                        return finish(
                            window,
                            "the GUI smoke window was not mapped on workspace 15",
                        )
                    focus_workspace(WORKSPACE)
                    focus_client(client)
                    sway_command(
                        f"[con_id={client['id']}] floating enable, "
                        "resize set width 760 height 900"
                    )
                except (OSError, RuntimeError, subprocess.SubprocessError) as error:
                    return finish(window, f"could not drive real desktop input: {error}")
                state["phase"] = "resize"
                state["checks"] = 0
                return GLib.SOURCE_CONTINUE

            state["checks"] = int(state["checks"]) + 1
            if state["phase"] == "resize":
                client = smoke_client(SMOKE_TEST_CLASS)
                if client is not None and client["size"] == [760, 900]:
                    desktop_input.type_text("c")
                    state["phase"] = "first-popout"
                    state["checks"] = 0
                    return GLib.SOURCE_CONTINUE
                if int(state["checks"]) < 30:
                    return GLib.SOURCE_CONTINUE
                return finish(window, "Sway did not resize the comment smoke window")
            if window._active_comment_id is None:
                if int(state["checks"]) < 30:
                    return GLib.SOURCE_CONTINUE
                return finish(window, "c did not focus the comment rail")
            rail_card = window._comment_cards[window._active_comment_id]
            rail_style = rail_card.get_style_context()
            popout = window._active_comment_float_card
            layer_parent = window._active_comment_layer.get_parent()
            if state["phase"] == "workspace":
                page = window.document._pages[0]
                if (
                    page.get_allocated_width() <= 1
                    or page.get_allocated_height() <= 1
                    or window.document._maximum_scroll() <= 0
                ):
                    if int(state["checks"]) < 30:
                        return GLib.SOURCE_CONTINUE
                    return finish(
                        window,
                        "workspace 15 did not allocate a scrollable rendered PDF",
                    )
                try:
                    client = smoke_client(SMOKE_TEST_CLASS)
                    if client is None:
                        return finish(
                            window,
                            "Sway lost the smoke-test window after focusing it",
                        )
                    drag_start, drag_end = _page_drag_points(page)
                    # GTK's Wayland backend reports these relative to the
                    # top-level surface because Wayland exposes no root
                    # window. Add Sway's compositor-space client origin.
                    window_x, window_y = client["at"]
                    page_local_x, page_local_y = (
                        page.get_window().get_root_coords(0, 0)
                    )
                    page_x = window_x + page_local_x
                    page_y = window_y + page_local_y
                    pointer_start = (
                        int(page_x + drag_start[0]),
                        int(page_y + drag_start[1]),
                    )
                    state["page_bounds"] = (
                        int(page_x),
                        int(page_y),
                        int(page_x + page.get_allocated_width()),
                        int(page_y + page.get_allocated_height()),
                    )
                    state["pointer_start"] = pointer_start
                    state["drag_end"] = (
                        int(page_x + drag_end[0]),
                        int(page_y + drag_end[1]),
                    )
                    desktop_input.move_cursor(*pointer_start)
                except (OSError, RuntimeError, subprocess.SubprocessError) as error:
                    return finish(window, f"could not drive real desktop input: {error}")
                state["phase"] = "pointer"
                state["checks"] = 0
                return GLib.SOURCE_CONTINUE
            if state["phase"] == "pointer":
                page = window.document._pages[0]
                if page._text_cursor is None:
                    if int(state["checks"]) < INPUT_SETTLE_CHECKS:
                        return GLib.SOURCE_CONTINUE
                    return finish(
                        window,
                        "real pointer input over the PDF did not restore its text cursor: "
                        f"target={state['pointer_start']}; "
                        f"actual={desktop_input.cursor_position()}; "
                        f"page_bounds={state['page_bounds']}; "
                        f"connector_has_window={window._comment_line_layer.get_has_window()}",
                    )
                desktop_input.left_button(pressed=True)
                desktop_input.move_cursor(*state["drag_end"])
                desktop_input.left_button(pressed=False)
                state["phase"] = "selection"
                state["checks"] = 0
                return GLib.SOURCE_CONTINUE
            if state["phase"] == "selection":
                page = window.document._pages[0]
                selection = page._text_selection
                if selection is None or not page.selected_text(selection).strip():
                    if int(state["checks"]) < INPUT_SETTLE_CHECKS:
                        return GLib.SOURCE_CONTINUE
                    return finish(
                        window,
                        "real pointer drag over the PDF did not select its text",
                    )
                adjustment = window.document.widget.get_vadjustment()
                state["scroll_before"] = adjustment.get_value()
                desktop_input.scroll_down()
                state["phase"] = "scroll"
                state["checks"] = 0
                return GLib.SOURCE_CONTINUE
            if state["phase"] == "scroll":
                adjustment = window.document.widget.get_vadjustment()
                if adjustment.get_value() <= float(state["scroll_before"]):
                    if int(state["checks"]) < INPUT_SETTLE_CHECKS:
                        return GLib.SOURCE_CONTINUE
                    return finish(
                        window,
                        "real wheel input over the PDF did not scroll the document",
                    )
                try:
                    desktop_input.key("Escape")
                except subprocess.SubprocessError as error:
                    return finish(window, f"could not send Esc through wtype: {error}")
                state["phase"] = "restored"
                state["checks"] = 0
                return GLib.SOURCE_CONTINUE
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
                rail_widths = {
                    card.get_allocated_width()
                    for card in window._comment_cards.values()
                }
                if len(rail_widths) != 1:
                    return finish(
                        window,
                        "comments in the rail use different widths: "
                        f"{sorted(rail_widths)}",
                    )
                if state["phase"] == "first-popout":
                    state["first_comment_id"] = window._active_comment_id
                    state["first_popout_width"] = popout_allocation.width
                    desktop_input.type_text("j")
                    state["phase"] = "second-popout"
                    state["checks"] = 0
                    return GLib.SOURCE_CONTINUE
                if state["phase"] == "second-popout":
                    if window._active_comment_id == state["first_comment_id"]:
                        if int(state["checks"]) < INPUT_SETTLE_CHECKS:
                            return GLib.SOURCE_CONTINUE
                        return finish(window, "j did not focus the long comment")
                    if popout_allocation.width != state["first_popout_width"]:
                        return finish(
                            window,
                            "short and long comments use different popout widths: "
                            f"{state['first_popout_width']}px and "
                            f"{popout_allocation.width}px",
                        )
                state["phase"] = "workspace"
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
