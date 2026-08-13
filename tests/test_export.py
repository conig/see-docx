from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from see_docx.viewer import (
    EXPORT_FORMATS,
    DocxWindow,
    Gdk,
    Gtk,
    _markdown_export_path,
    _pdf_export_path,
    _plain_text_export_path,
)


class _Panel:
    def __init__(self, visible: bool = False) -> None:
        self.visible = visible

    def get_visible(self) -> bool:
        return self.visible

    def hide(self) -> None:
        self.visible = False


class _Entry:
    def is_focus(self) -> bool:
        return False


class _ExportDestinationWindow:
    _choose_export_destination = DocxWindow._choose_export_destination

    def __init__(self) -> None:
        self.path = Path("/tmp/source.docx")
        self._export_panel = _Panel(visible=True)

    def _close_export(self) -> None:
        self._export_panel.hide()


class _MarkdownExportWindow:
    _activate_export_option = DocxWindow._activate_export_option

    def __init__(self) -> None:
        self._export_process = None
        self._export_index = EXPORT_FORMATS.index("Markdown")
        self.markdown_chosen = False

    def _choose_markdown_destination(self) -> None:
        self.markdown_chosen = True


class _ExportWindow:
    _on_key_press = DocxWindow._on_key_press

    def __init__(self) -> None:
        self._search_panel = _Panel()
        self._page_jump_panel = _Panel()
        self._outline_panel = _Panel()
        self._export_panel = _Panel()
        self._search_entry = _Entry()
        self._page_jump_entry = _Entry()
        self._url_hint_targets: dict[str, str] = {}
        self._pending_g = False
        self.actions: list[str] = []

    def _toggle_export(self) -> None:
        self.actions.append("toggle")
        self._export_panel.visible = not self._export_panel.visible

    def _move_export_selection(self, direction: int) -> None:
        self.actions.append(f"move:{direction}")

    def _activate_export_option(self) -> None:
        self.actions.append("activate")


class ExportTests(unittest.TestCase):
    def test_pdf_export_path_preserves_or_adds_the_pdf_suffix(self) -> None:
        self.assertEqual(_pdf_export_path(Path("report.pdf")), Path("report.pdf"))
        self.assertEqual(_pdf_export_path(Path("report.PDF")), Path("report.PDF"))
        self.assertEqual(_pdf_export_path(Path("report")), Path("report.pdf"))
        self.assertEqual(
            _pdf_export_path(Path("report.final")), Path("report.final.pdf")
        )

    def test_plain_text_is_an_export_option_with_a_txt_destination(self) -> None:
        self.assertIn("Plain text", EXPORT_FORMATS)
        self.assertEqual(
            _plain_text_export_path(Path("report.txt")), Path("report.txt")
        )
        self.assertEqual(
            _plain_text_export_path(Path("report.TXT")), Path("report.TXT")
        )
        self.assertEqual(
            _plain_text_export_path(Path("report")), Path("report.txt")
        )
        self.assertEqual(
            _plain_text_export_path(Path("report.final")), Path("report.final.txt")
        )

    def test_markdown_is_an_export_option_with_an_md_destination(self) -> None:
        self.assertIn("Markdown", EXPORT_FORMATS)
        self.assertEqual(
            _markdown_export_path(Path("report.md")), Path("report.md")
        )
        self.assertEqual(
            _markdown_export_path(Path("report.MD")), Path("report.MD")
        )
        self.assertEqual(
            _markdown_export_path(Path("report")), Path("report.md")
        )
        self.assertEqual(
            _markdown_export_path(Path("report.final")), Path("report.final.md")
        )

        window = _MarkdownExportWindow()
        window._activate_export_option()
        self.assertTrue(window.markdown_chosen)

    def test_export_tool_uses_e_then_j_k_and_enter(self) -> None:
        window = _ExportWindow()

        self.assertTrue(
            window._on_key_press(
                window, SimpleNamespace(keyval=Gdk.KEY_e, state=0)
            )
        )
        self.assertTrue(window._export_panel.visible)
        for key in (Gdk.KEY_j, Gdk.KEY_k, Gdk.KEY_Return):
            self.assertTrue(
                window._on_key_press(window, SimpleNamespace(keyval=key, state=0))
            )

        self.assertEqual(window.actions, ["toggle", "move:1", "move:-1", "activate"])

    def test_accepting_destination_closes_export_menu(self) -> None:
        # A completed chooser interaction starts the export and must also
        # dismiss the format menu, without requiring a follow-up Escape.
        window = _ExportDestinationWindow()
        destinations: list[Path] = []

        def start_export(destination: Path) -> bool:
            destinations.append(destination)
            return True

        dialog = Mock()
        dialog.run.return_value = Gtk.ResponseType.ACCEPT
        dialog.get_filename.return_value = "/tmp/exported"

        with (
            patch("see_docx.viewer.Gtk.FileChooserDialog", return_value=dialog),
            patch("see_docx.viewer.Gtk.FileFilter"),
        ):
            window._choose_export_destination(
                title="Export PDF",
                filename="source.pdf",
                filter_name="PDF documents",
                mime_type="application/pdf",
                pattern="*.pdf",
                export_path=_pdf_export_path,
                start_export=start_export,
            )

        self.assertEqual(destinations, [Path("/tmp/exported.pdf")])
        self.assertFalse(window._export_panel.visible)
