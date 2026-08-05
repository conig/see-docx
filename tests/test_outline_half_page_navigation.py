from __future__ import annotations

from types import SimpleNamespace
import unittest

from see_docx.viewer import DocxWindow, Gdk


class _Panel:
    def get_visible(self) -> bool:
        return True


class _Entry:
    def is_focus(self) -> bool:
        return False


class _Document:
    def __init__(self) -> None:
        self.scroll_commands: list[str] = []

    def scroll(self, command: str) -> None:
        self.scroll_commands.append(command)


class _Path:
    def __init__(self, value: int) -> None:
        self.value = value

    def to_string(self) -> str:
        return str(self.value)


class _OutlineTree:
    def __init__(self, *, focused: bool = False) -> None:
        self.focused = focused

    def get_visible_range(self) -> tuple[_Path, _Path]:
        # PyGObject returns the two paths directly, despite the underlying C
        # API also returning a success flag through its out parameters.
        return _Path(2), _Path(9)

    def is_focus(self) -> bool:
        return self.focused


class _OutlineRangeWindow:
    _outline_half_page_step = DocxWindow._outline_half_page_step

    def __init__(self) -> None:
        self._outline_tree = _OutlineTree()
        self._outline_row_paths = [_Path(index) for index in range(12)]

    def _visible_outline_indices(self) -> list[int]:
        return list(range(12))


class _OutlineHalfPageWindow:
    """Exercise the window's key-dispatch boundary with outline focus."""

    _on_key_press = DocxWindow._on_key_press

    def __init__(self) -> None:
        self._search_panel = _Panel()
        self._page_jump_panel = _Panel()
        self._outline_panel = _Panel()
        self._search_entry = _Entry()
        self._page_jump_entry = _Entry()
        self._outline_tree = _OutlineTree(focused=True)
        self._outline_count = 0
        self._pending_g = False
        self.document = _Document()
        self.outline_moves: list[int] = []

    def _outline_half_page_step(self) -> int:
        return 3

    def _move_outline_selection(self, direction: int) -> bool:
        self.outline_moves.append(direction)
        return True


class OutlineHalfPageNavigationTests(unittest.TestCase):
    def test_half_page_step_uses_the_real_gtk_visible_range_shape(self) -> None:
        window = _OutlineRangeWindow()

        self.assertEqual(window._outline_half_page_step(), 4)

    def test_ctrl_d_and_ctrl_u_move_half_a_visible_outline_page(self) -> None:
        # The outline tree owns focus while its menu is open, so Vim-style
        # half-page keys must move its selected heading, not the PDF viewport.
        window = _OutlineHalfPageWindow()
        control = Gdk.ModifierType.CONTROL_MASK

        self.assertTrue(
            window._on_key_press(
                window._outline_tree,
                SimpleNamespace(keyval=Gdk.KEY_d, state=control),
            )
        )
        self.assertTrue(
            window._on_key_press(
                window._outline_tree,
                SimpleNamespace(keyval=Gdk.KEY_u, state=control),
            )
        )

        self.assertEqual(window.outline_moves, [3, -3])
        self.assertEqual(window.document.scroll_commands, [])

    def test_window_key_propagation_respects_the_focused_outline_tree(self) -> None:
        # A tree key can bubble to the window, which must still use the
        # focused outline as the active navigation target.
        window = _OutlineHalfPageWindow()
        control = Gdk.ModifierType.CONTROL_MASK

        self.assertTrue(
            window._on_key_press(
                window,
                SimpleNamespace(keyval=Gdk.KEY_d, state=control),
            )
        )

        self.assertEqual(window.outline_moves, [3])
        self.assertEqual(window.document.scroll_commands, [])
