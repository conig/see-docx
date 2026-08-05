from __future__ import annotations

from types import SimpleNamespace
import unittest
from unittest.mock import patch

import cairo

from see_docx.viewer import (
    OUTLINE_LOCATOR_DURATION_MS,
    PdfPage,
    _outline_header_line_height,
    _outline_locator_frame,
)


class _RecordingContext:
    def __init__(self) -> None:
        self.rectangles: list[tuple[float, float, float, float]] = []
        self.arcs: list[tuple[float, float, float]] = []
        self.source_colors: list[tuple[float, float, float, float]] = []
        self.operators: list[int] = []

    def set_source_rgba(self, *components: float) -> None:
        self.source_colors.append(components)

    def set_operator(self, operator: int) -> None:
        self.operators.append(operator)

    def rectangle(self, x: float, y: float, width: float, height: float) -> None:
        self.rectangles.append((x, y, width, height))

    def arc(self, x: float, y: float, radius: float, *_angles: float) -> None:
        self.arcs.append((x, y, radius))

    def fill(self) -> None:
        pass

    def fill_preserve(self) -> None:
        pass

    def set_line_width(self, _width: float) -> None:
        pass

    def stroke(self) -> None:
        pass

    def save(self) -> None:
        pass

    def restore(self) -> None:
        pass

    def clip(self) -> None:
        pass


class OutlineLocatorAnimationTests(unittest.TestCase):
    def test_arrival_fills_then_releases_a_fading_heading_marker(self) -> None:
        arrival = _outline_locator_frame(0)
        settled = _outline_locator_frame(720)
        fading = _outline_locator_frame(1_400)
        finished = _outline_locator_frame(OUTLINE_LOCATOR_DURATION_MS)

        self.assertGreater(arrival.bloom_opacity, 0)
        self.assertGreater(settled.expansion, arrival.expansion)
        self.assertGreater(settled.expansion, 0.90)
        self.assertLess(fading.bloom_opacity, arrival.bloom_opacity)
        self.assertLess(fading.fill_opacity, settled.fill_opacity)
        self.assertEqual(finished.bloom_opacity, 0)
        self.assertEqual(finished.fill_opacity, 0)
        self.assertEqual(finished.anchor_opacity, 0)

    def test_fill_remains_visibly_in_progress_after_the_scroll_settles(self) -> None:
        # Navigation redraws after the scroll jump. A 200 ms sweep was
        # effectively complete before that redraw could register visually.
        self.assertLess(_outline_locator_frame(250).expansion, 0.60)

    def test_fill_is_anchored_to_the_bookmark_destination_not_layout_guess(self) -> None:
        # Text layout can point at body content; the visual arrival marker
        # must always paint at the bookmark location selected in the outline.
        page = PdfPage.__new__(PdfPage)
        page._outline_locator_top = 620.0
        page._height = 792.0
        page._width = 595.0
        page._zoom = 1.0
        early_context = _RecordingContext()
        late_context = _RecordingContext()
        with patch(
            "see_docx.viewer._theme_palette",
            return_value={"accent": "#5ced30", "highlight": "#eeffdb"},
        ):
            page._outline_locator_elapsed_ms = 180.0
            page._draw_outline_locator(early_context)
            page._outline_locator_elapsed_ms = 720.0
            page._draw_outline_locator(late_context)

        self.assertEqual(late_context.rectangles[0][1], 172.0)
        self.assertTrue(late_context.arcs)
        # The selected heading is the fixed ripple origin: only its radius
        # changes, while the square band clips the expanding circle.
        self.assertEqual({x for x, _y, _radius in early_context.arcs}, {72.0})
        self.assertEqual({x for x, _y, _radius in late_context.arcs}, {72.0})
        self.assertLess(early_context.arcs[0][2], late_context.arcs[0][2])
        self.assertLess(
            max(x + radius for x, _y, radius in late_context.arcs), page._width
        )

    def test_fill_height_matches_the_header_line_at_the_bookmark(self) -> None:
        # Bookmarks give us the destination, while Poppler's text layout gives
        # the actual rendered height of that particular header line.
        page = PdfPage.__new__(PdfPage)
        page._page = SimpleNamespace(
            get_size=lambda: (595.0, 792.0),
            get_text_layout=lambda: (
                True,
                [
                    SimpleNamespace(x1=80.0, y1=172.5, x2=90.0, y2=190.5),
                    SimpleNamespace(x1=90.0, y1=172.5, x2=99.0, y2=190.5),
                    SimpleNamespace(x1=72.0, y1=208.0, x2=78.0, y2=221.0),
                ],
            ),
        )
        page._outline_locator_top = 620.0
        page._height = 792.0
        page._width = 595.0
        page._zoom = 1.0
        context = _RecordingContext()

        with patch(
            "see_docx.viewer._theme_palette",
            return_value={"accent": "#5ced30", "highlight": "#eeffdb"},
        ):
            page._outline_locator_height = _outline_header_line_height(
                page._page, 620.0
            )
            page._outline_locator_elapsed_ms = 180.0
            page._draw_outline_locator(context)

        self.assertEqual(context.rectangles[0], (0.0, 172.0, 595.0, 18.0))

    def test_arrival_animation_uses_the_active_theme_colours(self) -> None:
        # The locator is drawn with Cairo, so its colours must be taken from
        # the GTK palette rather than retaining the old SC1 green hex values.
        page = PdfPage.__new__(PdfPage)
        page._outline_locator_top = 620.0
        page._outline_locator_elapsed_ms = 180.0
        page._height = 792.0
        page._width = 595.0
        page._zoom = 1.0
        context = _RecordingContext()

        with patch(
            "see_docx.viewer._theme_palette",
            return_value={"accent": "#123456", "highlight": "#abcdef"},
        ):
            page._draw_outline_locator(context)

        self.assertEqual(
            [colour[:3] for colour in context.source_colors],
            [
                (0x12 / 255, 0x34 / 255, 0x56 / 255),
                (0x12 / 255, 0x34 / 255, 0x56 / 255),
                (0xAB / 255, 0xCD / 255, 0xEF / 255),
            ],
        )
        self.assertEqual(context.operators, [cairo.OPERATOR_MULTIPLY])

    def test_search_highlight_uses_the_active_theme_colours(self) -> None:
        page = PdfPage.__new__(PdfPage)
        page._search_highlight = type(
            "Match",
            (),
            {"left": 10.0, "top": 40.0, "right": 30.0, "bottom": 20.0},
        )()
        page._height = 100.0
        page._zoom = 1.0
        context = _RecordingContext()

        with patch(
            "see_docx.viewer._theme_palette",
            return_value={"accent": "#123456", "highlight": "#abcdef"},
        ):
            page._draw_search_highlight(context)

        self.assertEqual(
            context.source_colors,
            [
                (0x12 / 255, 0x34 / 255, 0x56 / 255, 0.28),
                (0xAB / 255, 0xCD / 255, 0xEF / 255, 0.96),
            ],
        )
