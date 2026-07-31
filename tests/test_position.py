from __future__ import annotations

import unittest

from see_docx.position import (
    DocumentPosition,
    PageGeometry,
    capture_position,
    page_index_at_scroll,
    restore_position,
)


class PositionTests(unittest.TestCase):
    def test_restores_the_same_page_and_relative_offset(self) -> None:
        old_pages = [PageGeometry(0, 1000), PageGeometry(1040, 1000)]
        captured = capture_position(1290, old_pages, maximum_scroll=1500)

        self.assertEqual(captured.page_index, 1)
        self.assertAlmostEqual(captured.page_fraction or 0, 0.25)

        new_pages = [PageGeometry(0, 1100), PageGeometry(1140, 1400)]
        restored = restore_position(captured, new_pages, maximum_scroll=2100)
        self.assertAlmostEqual(restored, 1490)

    def test_falls_back_to_document_fraction_if_the_page_is_removed(self) -> None:
        captured = DocumentPosition(page_index=5, page_fraction=0.4, document_fraction=0.75)

        restored = restore_position(captured, [PageGeometry(0, 1000)], maximum_scroll=2000)

        self.assertEqual(restored, 1500)

    def test_never_restores_beyond_the_scrollable_range(self) -> None:
        captured = DocumentPosition(page_index=0, page_fraction=1.0, document_fraction=1.0)

        restored = restore_position(captured, [PageGeometry(0, 3000)], maximum_scroll=1800)

        self.assertEqual(restored, 1800)

    def test_position_in_page_gap_uses_document_fraction(self) -> None:
        pages = [PageGeometry(0, 100), PageGeometry(140, 100)]
        captured = capture_position(120, pages, maximum_scroll=200)

        self.assertIsNone(captured.page_index)
        self.assertEqual(restore_position(captured, pages, maximum_scroll=400), 240)

    def test_page_indicator_uses_the_next_page_in_a_gap(self) -> None:
        pages = [PageGeometry(24, 100), PageGeometry(142, 100), PageGeometry(260, 100)]

        self.assertEqual(page_index_at_scroll(0, pages, maximum_scroll=300), 0)
        self.assertEqual(page_index_at_scroll(130, pages, maximum_scroll=300), 1)
        self.assertEqual(page_index_at_scroll(250, pages, maximum_scroll=300), 2)

    def test_page_indicator_clamps_to_the_last_page(self) -> None:
        pages = [PageGeometry(0, 100), PageGeometry(120, 100)]

        self.assertEqual(page_index_at_scroll(900, pages, maximum_scroll=180), 1)
        self.assertIsNone(page_index_at_scroll(0, [], maximum_scroll=0))
