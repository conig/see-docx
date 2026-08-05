from __future__ import annotations

import unittest

from see_docx.viewer import _fit_zoom_for_viewport


class OutlineLayoutTests(unittest.TestCase):
    def test_fits_a_complete_page_with_margins_into_the_remaining_viewport(self) -> None:
        self.assertAlmostEqual(
            _fit_zoom_for_viewport(595.0, 792.0, 816.0, 728.0),
            min(744.0 / 595.0, 656.0 / 792.0),
        )
        self.assertAlmostEqual(_fit_zoom_for_viewport(595.0, 792.0, 667.0, 1000.0), 1.0)
        self.assertAlmostEqual(_fit_zoom_for_viewport(595.0, 792.0, 100.0, 100.0), 0.1)
