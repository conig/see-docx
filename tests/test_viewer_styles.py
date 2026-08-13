from __future__ import annotations

import unittest
from unittest.mock import patch

from see_docx.viewer import (
    COMMENTS_HEADER_HEIGHT,
    OUTLINE_NAV_SPACING,
    _app_css,
    _reading_progress_fraction,
    _reading_progress_frame,
)


class _Color:
    def __init__(self, value: str) -> None:
        self.value = value

    def to_string(self) -> str:
        return self.value


class _StyleContext:
    def lookup_color(self, name: str) -> tuple[bool, _Color]:
        colors = {
            "sc1-command-green": "#5ced30",
            "sc1-selection-fg": "#daff8c",
            "sc1-terran-blue": "#9db6e9",
        }
        if name in colors:
            return True, _Color(colors[name])
        return False, _Color("#000000")


class _Widget:
    def get_style_context(self) -> _StyleContext:
        return _StyleContext()


class ViewerStyleTests(unittest.TestCase):
    def _fallback_css(self) -> str:
        """Render the documented GTK fallback independent of local theme state."""

        with patch("see_docx.viewer._current_theme_state_roles", return_value=None):
            return _app_css(_Widget()).decode()

    def test_status_metadata_uses_the_standard_sc1_blue(self) -> None:
        css = self._fallback_css()

        self.assertIn(
            ".see-docx-page-indicator {\n"
            "  color: #9db6e9;\n",
            css,
        )
        self.assertIn(
            ".see-docx-path-status {\n"
            "  color: #9db6e9;\n",
            css,
        )
        self.assertIn(
            ".see-docx-status {\n"
            "  background-color: #202326;\n"
            "  padding: 6px 12px;\n"
            "}",
            css,
        )

    def test_reading_progress_uses_the_variant_accent(self) -> None:
        css = self._fallback_css()

        self.assertIn(
            ".see-docx-reading-progress progress {\n"
            "  min-height: 3px;\n"
            "  background-color: #5ced30;\n"
            "  border: 0;\n"
            "  border-radius: 0;\n"
            "}",
            css,
        )

    def test_action_notifications_use_the_active_theme_without_excess_chrome(self) -> None:
        css = self._fallback_css()

        self.assertIn(
            ".see-docx-notification {\n"
            "  background-color: #202326;\n"
            "  border: 1px solid #3b4148;\n"
            "  border-left: 3px solid #5ced30;\n"
            "  border-radius: 4px;\n"
            "  box-shadow: 0 8px 24px rgba(0, 0, 0, 0.34);\n",
            css,
        )
        self.assertIn(
            ".see-docx-notification-detail {\n"
            "  color: #a9b2bd;\n"
            "  font-size: 0.76em;\n",
            css,
        )

    def test_reading_progress_maps_and_clamps_scroll_position(self) -> None:
        self.assertEqual(_reading_progress_fraction(0, 400), 0.0)
        self.assertEqual(_reading_progress_fraction(100, 400), 0.25)
        self.assertEqual(_reading_progress_fraction(600, 400), 1.0)
        self.assertEqual(_reading_progress_fraction(-50, 400), 0.0)
        self.assertEqual(_reading_progress_fraction(0, 0), 0.0)

    def test_reading_progress_glides_to_its_latest_scroll_target(self) -> None:
        self.assertEqual(_reading_progress_frame(0.2, 0.8, 0), 0.2)
        self.assertAlmostEqual(_reading_progress_frame(0.2, 0.8, 90), 0.725)
        self.assertEqual(_reading_progress_frame(0.2, 0.8, 180), 0.8)
        self.assertEqual(_reading_progress_frame(0.2, 0.8, 360), 0.8)

    def test_outline_uses_variant_view_and_selection_foregrounds(self) -> None:
        css = self._fallback_css()

        self.assertIn(
            ".see-docx-outline treeview.view {\n"
            "  background-color: #202326;\n"
            "  color: #5ced30;\n"
            "  font-size: 0.96em;\n"
            "  -GtkTreeView-horizontal-separator: 7px;\n"
            "  -GtkTreeView-level-indentation: 7px;\n"
            "}",
            css,
        )

    def test_outline_uses_one_spacing_unit_around_tree_content(self) -> None:
        css = self._fallback_css()

        self.assertIn(
            ".see-docx-outline treeview.view {\n"
            "  background-color: #202326;\n"
            "  color: #5ced30;\n"
            "  font-size: 0.96em;\n"
            f"  -GtkTreeView-horizontal-separator: {OUTLINE_NAV_SPACING}px;\n"
            f"  -GtkTreeView-level-indentation: {OUTLINE_NAV_SPACING}px;\n"
            "}",
            css,
        )

    def test_active_outline_expanders_use_the_variant_highlight(self) -> None:
        css = self._fallback_css()

        self.assertIn(
            ".see-docx-outline treeview.view.expander:checked,\n"
            ".see-docx-outline treeview.view.expander:hover,\n"
            ".see-docx-outline treeview.view.expander:active {\n"
            "  color: #e7ebf3;\n"
            "}",
            css,
        )

    def test_search_uses_one_surface_without_an_inner_field_frame(self) -> None:
        css = self._fallback_css()

        self.assertIn(
            ".see-docx-search entry {\n"
            "  min-width: 0;\n"
            "  background-color: transparent;\n"
            "  border: 0;\n"
            "  box-shadow: none;\n",
            css,
        )

    def test_comment_header_uses_compact_hierarchy_and_subtle_key_hints(self) -> None:
        css = self._fallback_css()

        self.assertIn(
            ".see-docx-comments-header {\n"
            f"  min-height: {COMMENTS_HEADER_HEIGHT}px;\n"
            "  padding: 4px 12px;\n"
            "}",
            css,
        )
        self.assertIn(
            ".see-docx-comments-header-top {\n"
            "  min-height: 16px;\n"
            "}",
            css,
        )
        self.assertIn(
            ".see-docx-comments-key {\n"
            "  color: #a9b2bd;\n"
            "  font-size: 0.68em;\n",
            css,
        )
        self.assertIn(
            ".see-docx-comment-thread.comment-rail-ghost {\n"
            "  background-color: transparent;\n"
            "  border: 1px dashed #3b4148;\n"
            "  border-left: 2px dashed #5ced30;\n"
            "  border-radius: 5px;\n"
            "  box-shadow: 0 7px 18px rgba(0, 0, 0, 0.24),\n",
            css,
        )

    def test_persistent_search_state_uses_the_variant_highlight(self) -> None:
        css = self._fallback_css()

        self.assertIn(
            ".see-docx-search-session {\n"
            "  color: #e7ebf3;\n"
            "  font-size: 0.78em;\n"
            "  font-weight: 700;\n"
            "}",
            css,
        )
        self.assertIn(
            ".see-docx-outline treeview.view:selected,\n"
            ".see-docx-outline treeview.view:selected:focus {\n"
            "  box-shadow: none;\n"
            "  color: #daff8c;\n"
            "}",
            css,
        )
