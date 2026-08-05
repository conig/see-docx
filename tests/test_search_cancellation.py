from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import Mock, patch

from see_docx.viewer import DocxWindow, Gdk


class _Panel:
    def __init__(self, visible: bool = False) -> None:
        self.visible = visible

    def get_visible(self) -> bool:
        return self.visible

    def hide(self) -> None:
        self.visible = False


class _Entry:
    def __init__(self, text: str) -> None:
        self.text = text

    def set_text(self, text: str) -> None:
        self.text = text

    def is_focus(self) -> bool:
        return False


class _Widget:
    def grab_focus(self) -> None:
        pass


class _Page:
    def __init__(self) -> None:
        self._search_highlight: object | None = object()


class _Document:
    def __init__(self) -> None:
        self.widget = _Widget()
        self._pages = [_Page()]
        self.copy_all_requested = False

    def clear_search_highlight(self) -> None:
        for page in self._pages:
            page._search_highlight = None

    def copy_all_text(self) -> bool:
        self.copy_all_requested = True
        return True


class _SearchWindow:
    """The real key handlers with only their GTK display objects replaced."""

    _toggle_search = DocxWindow._toggle_search
    _cancel_search = DocxWindow._cancel_search
    _clear_search_session_status = DocxWindow._clear_search_session_status
    _queue_search_match_marker_redraw = DocxWindow._queue_search_match_marker_redraw
    _on_search_key_press = DocxWindow._on_search_key_press
    _on_key_press = DocxWindow._on_key_press
    _copy_path = DocxWindow._copy_path
    _copy_all_text = DocxWindow._copy_all_text

    def __init__(self) -> None:
        self._search_panel = _Panel(visible=True)
        self._page_jump_panel = _Panel()
        self._outline_panel = _Panel()
        self._search_entry = _Entry("purpose")
        self._search_matches = [object(), object()]
        self._search_index = 1
        self._search_status = _Entry("2/2")
        self.path = Path("/tmp/report.docx")
        self.document = _Document()
        self._pending_g = False

    def _move_search_match(self, _direction: int) -> bool:
        return bool(self._search_matches)


class SearchCancellationTests(unittest.TestCase):
    def test_a_copies_all_document_text(self) -> None:
        window = _SearchWindow()

        self.assertTrue(
            window._on_key_press(
                window, SimpleNamespace(keyval=Gdk.KEY_a, state=0)
            )
        )

        self.assertTrue(window.document.copy_all_requested)

    def test_y_copies_the_resolved_local_document_path(self) -> None:
        window = _SearchWindow()
        clipboard = Mock()

        with patch("see_docx.viewer.Gtk.Clipboard.get", return_value=clipboard):
            self.assertTrue(
                window._on_key_press(
                    window, SimpleNamespace(keyval=Gdk.KEY_y, state=0)
                )
            )

        clipboard.set_text.assert_called_once_with("/tmp/report.docx", -1)

    def test_escape_clears_a_committed_search_and_disables_result_navigation(self) -> None:
        # Mirrors Escape after Enter has closed the prompt but retained the
        # highlighted result for n/N navigation.
        window = _SearchWindow()

        self.assertTrue(
            window._on_search_key_press(
                window._search_entry, SimpleNamespace(keyval=Gdk.KEY_Escape, state=0)
            )
        )
        self.assertFalse(window._search_panel.get_visible())
        self.assertEqual(window._search_entry.text, "")
        self.assertEqual(window._search_matches, [])
        self.assertEqual(window._search_index, -1)
        self.assertEqual(window._search_status.text, "")
        self.assertTrue(
            all(page._search_highlight is None for page in window.document._pages)
        )
        self.assertFalse(
            window._on_key_press(window, SimpleNamespace(keyval=Gdk.KEY_n, state=0))
        )
        self.assertFalse(
            window._on_key_press(window, SimpleNamespace(keyval=Gdk.KEY_N, state=0))
        )
