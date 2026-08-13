#!/usr/bin/env python3
"""Exercise an actual GTK pointer-selection drag across a PDF page break."""

from __future__ import annotations

import argparse
from pathlib import Path
import tempfile
from types import SimpleNamespace
import zipfile

import gi

gi.require_version("Gdk", "3.0")
gi.require_version("Gtk", "3.0")

from gi.repository import Gdk, Gio, GLib, Gtk

from see_docx.viewer import APPLICATION_ID, DocxWindow, PdfPage, TextSelection


TIMEOUT_MS = 20_000
SMOKE_TEST_CLASS = "codex-smoke-test"
HEADER_TEXT = "SELECTION SMOKE HEADER"
FOOTER_TEXT = "SELECTION SMOKE FOOTER"
_OFFICE_RELATIONSHIPS = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
)
_WORDPROCESSINGML = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"


def _add_running_matter(source: Path, destination: Path) -> None:
    """Give an existing DOCX distinctive repeated header and footer text."""

    with zipfile.ZipFile(source) as archive:
        members = {name: archive.read(name) for name in archive.namelist()}

    document_references: list[bytes] = []
    relationship_entries: list[bytes] = []
    content_type_entries: list[bytes] = []
    for flow, relationship_id, visible_text in (
        ("header", "rIdSelectionSmokeHeader", HEADER_TEXT),
        ("footer", "rIdSelectionSmokeFooter", FOOTER_TEXT),
    ):
        document_references.append(
            f'<w:{flow}Reference w:type="default" r:id="{relationship_id}"/>'.encode()
        )
        relationship_entries.append(
            (
                f'<Relationship Id="{relationship_id}" '
                f'Type="{_OFFICE_RELATIONSHIPS}/{flow}" '
                f'Target="{flow}-selection-smoke.xml"/>'
            ).encode()
        )
        content_type_entries.append(
            (
                f'<Override PartName="/word/{flow}-selection-smoke.xml" '
                "ContentType=\"application/vnd.openxmlformats-officedocument."
                f'wordprocessingml.{flow}+xml\"/>'
            ).encode()
        )
        root_name = "hdr" if flow == "header" else "ftr"
        members[f"word/{flow}-selection-smoke.xml"] = (
            '<?xml version="1.0" encoding="UTF-8"?>'
            f'<w:{root_name} xmlns:w="{_WORDPROCESSINGML}">'
            f"<w:p><w:r><w:t>{visible_text}</w:t></w:r></w:p>"
            f"</w:{root_name}>"
        ).encode()

    document_xml = members["word/document.xml"]
    if b"<w:sectPr>" not in document_xml:
        raise RuntimeError("fixture has no conventional Word section properties")
    members["word/document.xml"] = document_xml.replace(
        b"<w:sectPr>", b"<w:sectPr>" + b"".join(document_references), 1
    )
    members["word/_rels/document.xml.rels"] = members[
        "word/_rels/document.xml.rels"
    ].replace(b"</Relationships>", b"".join(relationship_entries) + b"</Relationships>")
    members["[Content_Types].xml"] = members["[Content_Types].xml"].replace(
        b"</Types>", b"".join(content_type_entries) + b"</Types>"
    )
    with zipfile.ZipFile(destination, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, content in members.items():
            archive.writestr(name, content)


def _flow_endpoint(page: PdfPage, flow: str, *, last: bool) -> tuple[float, float]:
    """Return the first or last rendered glyph centre in one selection flow."""

    text = page._page.get_text()
    has_layout, rectangles = page._page.get_text_layout()
    if not has_layout or len(text) != len(rectangles):
        raise RuntimeError("rendered page has no aligned glyph layout")
    indices = [
        index
        for index, character in enumerate(text)
        if character not in "\r\n"
        and page._selection_flow_map.get(index, "main") == flow
    ]
    if not indices:
        raise RuntimeError(f"rendered page has no {flow} glyphs")
    rectangle = rectangles[indices[-1 if last else 0]]
    return (
        (float(rectangle.x1) + float(rectangle.x2)) / 2,
        (float(rectangle.y1) + float(rectangle.y2)) / 2,
    )


def _verify_rendered_selection_flows(window: DocxWindow) -> str | None:
    """Check real Poppler glyph selection without requiring scroll allocation."""

    first_page, second_page = window.document._pages[:2]

    def whole_page(page: PdfPage) -> TextSelection:
        return TextSelection(0.0, 0.0, page._width, page._height)

    first_page.set_text_selection(
        whole_page(first_page),
        start=_flow_endpoint(first_page, "main", last=False),
        end=(first_page._width, first_page._height),
        flow="main",
    )
    second_page.set_text_selection(
        whole_page(second_page),
        start=(0.0, 0.0),
        end=_flow_endpoint(second_page, "main", last=True),
        flow="main",
    )
    body_segments = (
        first_page.selected_text(first_page._text_selection),
        second_page.selected_text(second_page._text_selection),
    )
    if any(
        HEADER_TEXT in segment or FOOTER_TEXT in segment
        for segment in body_segments
    ):
        return "main-text highlight crossed into a repeated header or footer"

    for flow, expected in (("header", HEADER_TEXT), ("footer", FOOTER_TEXT)):
        first_page.set_text_selection(
            whole_page(first_page),
            start=_flow_endpoint(first_page, flow, last=False),
            end=(first_page._width, first_page._height),
            flow=flow,
        )
        if first_page.selected_text(first_page._text_selection) != expected:
            return f"{flow} selection escaped its rendered story"
    return None


def main() -> int:
    GLib.set_prgname(SMOKE_TEST_CLASS)
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path, help="A multi-page DOCX fixture")
    parser.add_argument(
        "--selection-only",
        action="store_true",
        help="Verify rendered selection flows without the scrolling checks",
    )
    arguments = parser.parse_args()
    source = arguments.path.resolve()
    if not source.is_file():
        parser.error(f"No such DOCX: {source}")

    with tempfile.TemporaryDirectory(prefix="see-docx-selection-smoke-") as directory:
        watched = Path(directory) / "watched.docx"
        _add_running_matter(source, watched)
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
            if arguments.selection_only:
                try:
                    error = _verify_rendered_selection_flows(window)
                except RuntimeError as runtime_error:
                    error = str(runtime_error)
                if error:
                    return finish(window, error)
                print("rendered body, header, and footer selection flows stay separate")
                return finish(window)
            # This must inspect the realized widget tree: calling PdfPage's
            # handlers directly would bypass an invisible overlay that steals
            # the user's real mouse drag and wheel events.
            comment_layer_parent = window._active_comment_layer.get_parent()
            if (
                window._active_comment_layer.get_visible()
                and isinstance(comment_layer_parent, Gtk.Overlay)
                and not comment_layer_parent.get_overlay_pass_through(
                    window._active_comment_layer
                )
            ):
                return finish(
                    window,
                    "the inactive comment overlay intercepts document pointer input",
                )
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
                visible_segments = (
                    first_page.selected_text(first_page._text_selection),
                    second_page.selected_text(second_page._text_selection),
                )
                if any(
                    HEADER_TEXT in segment or FOOTER_TEXT in segment
                    for segment in visible_segments
                ):
                    return finish(
                        window,
                        "main-text highlight crossed into a repeated header or footer",
                    )
                selected = Gtk.Clipboard.get(
                    Gdk.SELECTION_CLIPBOARD
                ).wait_for_text()
                if selected is None:
                    return finish(window, "cross-page drag did not publish clipboard text")
                if HEADER_TEXT in selected or FOOTER_TEXT in selected:
                    return finish(
                        window,
                        "main-text selection crossed into a repeated header or footer",
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
                60_000 if arguments.selection_only else TIMEOUT_MS,
                lambda: finish(
                    window,
                    "timed out waiting for a rendered multi-page preview "
                    f"({window._last_status}; pages={window.document.page_count}; "
                    f"max-scroll={window.document._maximum_scroll()}; state={state})",
                ),
            )

        application.connect("activate", activate)
        application.run(["see-docx-selection-smoke"])
        if error := result.get("error"):
            print(f"FAIL: {error}")
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
