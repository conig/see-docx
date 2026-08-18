from __future__ import annotations

import unittest
from unittest.mock import patch

from see_docx.viewer import (
    ACTION_NOTIFICATION_DURATION_MS,
    DocxWindow,
    GLib,
)


class _Style:
    def __init__(self) -> None:
        self.classes: set[str] = set()

    def add_class(self, name: str) -> None:
        self.classes.add(name)

    def remove_class(self, name: str) -> None:
        self.classes.discard(name)


class _Box:
    def __init__(self) -> None:
        self.style = _Style()

    def get_style_context(self) -> _Style:
        return self.style


class _Label:
    def __init__(self) -> None:
        self.text = ""
        self.visible = True

    def set_text(self, text: str) -> None:
        self.text = text

    def set_visible(self, visible: bool) -> None:
        self.visible = visible


class _Revealer:
    def __init__(self) -> None:
        self.revealed = False

    def set_reveal_child(self, revealed: bool) -> None:
        self.revealed = revealed


class ActionNotificationTests(unittest.TestCase):
    def _window(self) -> DocxWindow:
        window = DocxWindow.__new__(DocxWindow)
        window._notification_source = 41
        window._notification_box = _Box()
        window._notification_mark = _Label()
        window._notification_title = _Label()
        window._notification_detail = _Label()
        window._notification_revealer = _Revealer()
        return window

    def test_new_action_replaces_the_previous_notification_and_timer(self) -> None:
        window = self._window()

        with (
            patch("see_docx.viewer.GLib.source_remove") as remove,
            patch("see_docx.viewer.GLib.timeout_add", return_value=73) as timeout,
        ):
            DocxWindow._show_notification(
                window,
                "Path copied",
                "~/reports/plan.docx",
            )

        remove.assert_called_once_with(41)
        timeout.assert_called_once_with(
            ACTION_NOTIFICATION_DURATION_MS,
            window._hide_notification,
        )
        self.assertEqual(window._notification_source, 73)
        self.assertTrue(window._notification_revealer.revealed)
        self.assertEqual(window._notification_mark.text, "✓")
        self.assertEqual(window._notification_title.text, "Path copied")
        self.assertEqual(
            window._notification_detail.text,
            "~/reports/plan.docx",
        )
        self.assertTrue(window._notification_detail.visible)
        self.assertNotIn("warning", window._notification_box.style.classes)

    def test_warning_and_dismissal_have_distinct_state(self) -> None:
        window = self._window()
        window._notification_source = 0

        with patch("see_docx.viewer.GLib.timeout_add", return_value=73):
            DocxWindow._show_notification(
                window,
                "Nothing to copy",
                success=False,
            )

        self.assertEqual(window._notification_mark.text, "!")
        self.assertIn("warning", window._notification_box.style.classes)
        self.assertFalse(window._notification_detail.visible)

        self.assertEqual(DocxWindow._hide_notification(window), GLib.SOURCE_REMOVE)
        self.assertEqual(window._notification_source, 0)
        self.assertFalse(window._notification_revealer.revealed)

    def test_whole_table_copy_uses_the_standard_success_notification(self) -> None:
        window = self._window()
        window._notification_source = 0

        with patch("see_docx.viewer.GLib.timeout_add", return_value=73):
            DocxWindow._on_table_copied(window)

        self.assertTrue(window._notification_revealer.revealed)
        self.assertEqual(window._notification_title.text, "Table copied")
        self.assertEqual(
            window._notification_detail.text,
            "The complete table is ready to paste",
        )


if __name__ == "__main__":
    unittest.main()
