from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
import unittest

from see_docx.viewer import (
    EXPORT_FORMATS,
    DocxWindow,
    Gdk,
    _pdf_export_path,
    _plain_text_export_path,
)


class _Panel:
    def __init__(self, visible: bool = False) -> None:
        self.visible = visible

    def get_visible(self) -> bool:
        return self.visible


class _Entry:
    def is_focus(self) -> bool:
        return False


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
