from __future__ import annotations

import unittest

from see_docx.viewer import _contextual_scroll_target, _search_scroll_target


class OutlineContextTests(unittest.TestCase):
    def test_centers_a_search_result_in_the_viewport_and_clamps(self) -> None:
        self.assertEqual(
            _search_scroll_target(destination=900, viewport_height=600, maximum_scroll=2000),
            600,
        )
        self.assertEqual(
            _search_scroll_target(destination=100, viewport_height=600, maximum_scroll=2000),
            0,
        )
        self.assertEqual(
            _search_scroll_target(destination=3000, viewport_height=600, maximum_scroll=2000),
            2000,
        )

    def test_places_a_heading_one_third_down_the_viewport_and_clamps(self) -> None:
        self.assertEqual(
            _contextual_scroll_target(destination=900, viewport_height=600, maximum_scroll=2000),
            720,
        )
        self.assertEqual(
            _contextual_scroll_target(destination=100, viewport_height=600, maximum_scroll=2000),
            0,
        )
        self.assertEqual(
            _contextual_scroll_target(destination=3000, viewport_height=600, maximum_scroll=2000),
            2000,
        )
