from __future__ import annotations

import unittest
from types import SimpleNamespace
from unittest.mock import Mock, patch

import see_docx.viewer as viewer
from see_docx.viewer import PdfDocumentView, PdfPage, TextSelection, _text_selection_bounds


class TextSelectionTests(unittest.TestCase):
    @staticmethod
    def _selection_page(
        lines: tuple[tuple[str, float, str], ...], allocation_y: float
    ) -> PdfPage:
        """Build one laid-out PDF page with semantic header/body/footer glyphs."""

        text_parts: list[str] = []
        rectangles: list[SimpleNamespace] = []
        selection_flows: dict[int, str] = {}
        for line_index, (line, y, flow) in enumerate(lines):
            if line_index:
                text_parts.append("\n")
                rectangles.append(SimpleNamespace(x1=0.0, y1=y, x2=0.0, y2=y))
            for character_index, character in enumerate(line):
                index = sum(len(part) for part in text_parts)
                text_parts.append(character)
                rectangles.append(
                    SimpleNamespace(
                        x1=character_index * 5.0,
                        y1=y,
                        x2=(character_index + 1) * 5.0,
                        y2=y + 4.0,
                    )
                )
                selection_flows[index] = flow

        text = "".join(text_parts)
        poppler_page = Mock()
        poppler_page.get_text.return_value = text
        poppler_page.get_text_layout.return_value = (True, rectangles)
        page = PdfPage.__new__(PdfPage)
        page._page = poppler_page
        page._width = 200.0
        page._height = 200.0
        page._zoom = 1.0
        page._text_selection = None
        page._text_selection_start = None
        page._text_selection_end = None
        page._selection_flow_map = selection_flows
        page._source_character_map = {}
        page._copy_selection = Mock()
        page.queue_draw = Mock()
        page.get_allocation = Mock(
            return_value=SimpleNamespace(
                x=0.0,
                y=allocation_y,
                width=200.0,
                height=200.0,
            )
        )
        return page

    @staticmethod
    def _selection_view(pages: list[PdfPage]) -> PdfDocumentView:
        view = PdfDocumentView.__new__(PdfDocumentView)
        view._pages = pages
        view._rich_source = None
        view._source_text = None
        view._selection_anchor = None
        view._selection_endpoint = None
        view._selection_anchor_page = None
        view._selection_drag_active = False
        view._selection_auto_scroll_source = 0
        view._selection_auto_scroll_page = None
        view._selection_auto_scroll_point = None
        return view

    def test_body_selection_crosses_pages_without_headers_or_footers(self) -> None:
        # A document drag may span body text on adjacent pages, but repeated
        # running matter between those endpoints is not part of the main flow.
        first = self._selection_page(
            (
                ("HEADER ONE", 10.0, "header"),
                ("Body first", 80.0, "main"),
                ("FOOTER ONE", 150.0, "footer"),
            ),
            0.0,
        )
        second = self._selection_page(
            (
                ("HEADER TWO", 10.0, "header"),
                ("Body second", 80.0, "main"),
                ("FOOTER TWO", 150.0, "footer"),
            ),
            228.0,
        )
        view = self._selection_view([first, second])

        view._on_page_selection_event("begin", first, (0.0, 82.0))
        view._on_page_selection_event("end", second, (55.0, 82.0))

        first._copy_selection.assert_called_once_with("Body first\nBody second")

    def test_header_and_footer_selections_cannot_escape_their_starting_flow(
        self,
    ) -> None:
        # Header/footer text is independently selectable, but neither flow may
        # grow into body text or continue onto another page.
        for flow, start_y, expected in (
            ("header", 12.0, "HEADER ONE"),
            ("footer", 152.0, "FOOTER ONE"),
        ):
            with self.subTest(flow=flow):
                first = self._selection_page(
                    (
                        ("HEADER ONE", 10.0, "header"),
                        ("Body first", 80.0, "main"),
                        ("FOOTER ONE", 150.0, "footer"),
                    ),
                    0.0,
                )
                second = self._selection_page(
                    (
                        ("HEADER TWO", 10.0, "header"),
                        ("Body second", 80.0, "main"),
                        ("FOOTER TWO", 150.0, "footer"),
                    ),
                    228.0,
                )
                view = self._selection_view([first, second])

                view._on_page_selection_event("begin", first, (0.0, start_y))
                view._on_page_selection_event("end", second, (55.0, 82.0))

                first._copy_selection.assert_called_once_with(expected)

    def test_dynamic_page_number_stays_in_its_static_header_flow(self) -> None:
        # OOXML stores a cached PAGE result (for example, "1"), while each
        # rendered page has a different number. Static story text must locate
        # the line and classify the dynamic glyph beside it as header text.
        text = "Page 2\nBody text"
        rectangles = [
            SimpleNamespace(
                x1=index * 5.0,
                y1=10.0 if index < 6 else 80.0,
                x2=(index + 1) * 5.0,
                y2=14.0 if index < 6 else 84.0,
            )
            for index in range(len(text))
        ]

        flow_map = viewer._rendered_selection_flow_map(
            text,
            rectangles,
            200.0,
            {"header": ("Page 1", "Page ")},
        )

        self.assertEqual(
            {flow_map[index] for index in range(len("Page 2"))},
            {"header"},
        )
        self.assertNotIn(len("Page 2\n"), flow_map)

    def test_copy_all_text_prefers_paragraph_preserving_docx_source(self) -> None:
        page = SimpleNamespace(_copy_selection=Mock())
        view = PdfDocumentView.__new__(PdfDocumentView)
        view._source_text = "First paragraph.\nSecond paragraph."
        view._pages = [page]

        self.assertTrue(view.copy_all_text())

        page._copy_selection.assert_called_once_with(
            "First paragraph.\nSecond paragraph."
        )

    def test_horizontal_drag_expands_to_a_selectable_text_line(self) -> None:
        selection = _text_selection_bounds(
            start=(120.0, 240.0),
            end=(310.0, 240.0),
            page_width=595.0,
            page_height=792.0,
            line_padding=4.0,
        )

        self.assertEqual(
            selection,
            TextSelection(left=120.0, top=236.0, right=310.0, bottom=244.0),
        )

    def test_drag_bounds_stay_inside_the_pdf_page(self) -> None:
        selection = _text_selection_bounds(
            start=(-30.0, -10.0),
            end=(640.0, 810.0),
            page_width=595.0,
            page_height=792.0,
            line_padding=8.0,
        )

        self.assertEqual(
            selection,
            TextSelection(left=0.0, top=0.0, right=595.0, bottom=792.0),
        )

    def test_drag_extracts_glyph_text_and_updates_the_visible_selection(self) -> None:
        poppler_page = Mock()
        # Poppler inserts a newline at every visual line wrap. Copying this
        # single Writer paragraph must preserve its reflowable text instead.
        poppler_page.get_selected_text.return_value = "selected\ntext"
        page = PdfPage.__new__(PdfPage)
        page._selection_anchor = (120.0, 240.0)
        page._width = 595.0
        page._height = 792.0
        page._zoom = 1.0
        page._page = poppler_page
        page._copy_selection = Mock()
        page.queue_draw = Mock()

        self.assertTrue(page._update_text_selection((310.0, 240.0)))
        self.assertEqual(
            page._text_selection,
            TextSelection(left=120.0, top=236.0, right=310.0, bottom=244.0),
        )
        page._copy_selection.assert_called_once_with("selected text")
        page.queue_draw.assert_called_once_with()

    def test_drag_joins_visual_wraps_before_copying(self) -> None:
        poppler_page = Mock()
        poppler_page.get_selected_text.return_value = "wrapped\nparagraph"
        page = PdfPage.__new__(PdfPage)
        page._selection_anchor = (120.0, 240.0)
        page._width = 595.0
        page._height = 792.0
        page._zoom = 1.0
        page._page = poppler_page
        page._copy_selection = Mock()
        page.queue_draw = Mock()

        page._update_text_selection((310.0, 240.0))

        page._copy_selection.assert_called_once_with("wrapped paragraph")

    def test_horizontal_drag_uses_glyph_layout_to_exclude_the_next_row(self) -> None:
        # Poppler's rectangular text extractor can claim the following visual
        # line. Its text layout is aligned with the rendered page instead.
        poppler_page = Mock()
        poppler_page.get_text.return_value = "first\nsecond"
        poppler_page.get_text_layout.return_value = (
            True,
            [
                *(SimpleNamespace(x1=index * 5.0, y1=8.0, x2=(index + 1) * 5.0, y2=12.0) for index in range(5)),
                SimpleNamespace(x1=25.0, y1=12.0, x2=25.0, y2=12.0),
                *(SimpleNamespace(x1=index * 5.0, y1=28.0, x2=(index + 1) * 5.0, y2=32.0) for index in range(6)),
            ],
        )
        poppler_page.get_selected_text.return_value = "first\nsecond"
        page = PdfPage.__new__(PdfPage)
        page._selection_anchor = (0.0, 10.0)
        page._width = 595.0
        page._height = 792.0
        page._zoom = 1.0
        page._page = poppler_page
        page._copy_selection = Mock()
        page.queue_draw = Mock()

        page._update_text_selection((25.0, 10.0))

        page._copy_selection.assert_called_once_with("first")

    def test_drag_selects_an_ordered_text_range_not_a_rectangle(self) -> None:
        # A normal text drag begins part way through the first row, includes
        # complete rows in between, then stops part way through the last row.
        # It must not behave like a rectangular marquee.
        text = "first line\nsecond line"
        rectangles = []
        x, y = 0.0, 8.0
        for character in text:
            if character == "\n":
                rectangles.append(
                    SimpleNamespace(x1=x, y1=y + 4.0, x2=x, y2=y + 4.0)
                )
                x, y = 0.0, y + 20.0
                continue
            rectangles.append(
                SimpleNamespace(x1=x, y1=y, x2=x + 5.0, y2=y + 4.0)
            )
            x += 5.0
        poppler_page = Mock()
        poppler_page.get_text.return_value = text
        poppler_page.get_text_layout.return_value = (True, rectangles)
        page = PdfPage.__new__(PdfPage)
        page._selection_anchor = (12.5, 10.0)
        page._width = 595.0
        page._height = 792.0
        page._zoom = 1.0
        page._page = poppler_page
        page._copy_selection = Mock()
        page.queue_draw = Mock()

        page._update_text_selection((17.5, 30.0))

        page._copy_selection.assert_called_once_with("rst line seco")
        selected, selected_rectangles = page._layout_text_selection(
            page._text_selection
        )
        self.assertEqual(selected, "rst line\nseco")
        rows: dict[float, list[SimpleNamespace]] = {}
        for rectangle in selected_rectangles:
            rows.setdefault(rectangle.y1, []).append(rectangle)
        self.assertEqual(
            [
                (
                    min(rectangle.x1 for rectangle in row),
                    max(rectangle.x2 for rectangle in row),
                )
                for _y, row in sorted(rows.items())
            ],
            [(10.0, 50.0), (0.0, 20.0)],
        )

    def test_copy_restores_docx_paragraph_boundaries_hidden_by_pdf_wraps(self) -> None:
        restore = getattr(viewer, "_restore_docx_paragraph_boundaries", None)
        self.assertIsNotNone(restore)

        self.assertEqual(
            restore(
                "First paragraph ends here.\nSecond paragraph begins here.",
                "First paragraph ends here. Second paragraph begins here.",
            ),
            "First paragraph ends here.\nSecond paragraph begins here.",
        )

    def test_pointer_motion_uses_the_text_selection_cursor(self) -> None:
        # Hovering a rendered PDF page must advertise that its glyphs can be
        # selected, even before the reader starts a drag.
        page = PdfPage.__new__(PdfPage)
        page._selection_handler = Mock()
        page._set_text_cursor = Mock()

        self.assertTrue(page._on_motion(page, SimpleNamespace(x=120.0, y=240.0)))

        page._set_text_cursor.assert_called_once_with()

    def test_copy_selection_publishes_to_regular_and_primary_clipboards(self) -> None:
        page = PdfPage.__new__(PdfPage)
        regular = Mock()
        primary = Mock()

        with patch(
            "see_docx.viewer.Gtk.Clipboard.get", side_effect=(regular, primary)
        ):
            page._copy_selection("selected text")

        regular.set_text.assert_called_once_with("selected text", -1)
        primary.set_text.assert_called_once_with("selected text", -1)
