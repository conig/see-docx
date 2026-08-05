from __future__ import annotations

from types import SimpleNamespace
import unittest
from unittest.mock import Mock

from see_docx.viewer import Gdk, PdfDocumentView, PdfPage, TextSelection


class _Adjustment:
    def __init__(self, value: float) -> None:
        self.value = value

    def get_value(self) -> float:
        return self.value


class TextSelectionScrollingTests(unittest.TestCase):
    @staticmethod
    def _page(
        *, top: float, selected_text: str
    ) -> tuple[PdfPage, Mock]:
        poppler_page = Mock()
        poppler_page.get_selected_text.return_value = selected_text
        page = PdfPage.__new__(PdfPage)
        page._text_selection = None
        page._width = 595.0
        page._height = 792.0
        page._zoom = 1.0
        page._page = poppler_page
        page.queue_draw = Mock()
        page.get_allocation = Mock(
            return_value=SimpleNamespace(x=0.0, y=top, height=792.0)
        )
        return page, poppler_page

    def test_active_drag_follows_a_viewport_scroll(self) -> None:
        """A wheel scroll during a held drag must extend the same selection.

        The pointer remains at y=240 in the viewport while the document moves
        down by 100 pixels, so the selected page coordinate must move to 340.
        """

        poppler_page = Mock()
        poppler_page.get_selected_text.return_value = "selected text"
        page = PdfPage.__new__(PdfPage)
        page._selection_anchor = (120.0, 240.0)
        page._selection_pointer = (310.0, 240.0)
        page._width = 595.0
        page._height = 792.0
        page._zoom = 1.0
        page._page = poppler_page
        page._copy_selection = Mock()
        page.queue_draw = Mock()
        page.get_allocation = Mock(
            return_value=SimpleNamespace(x=0.0, y=0.0, height=792.0)
        )
        page._update_text_selection(page._selection_pointer)

        view = PdfDocumentView.__new__(PdfDocumentView)
        view._pages = [page]
        view._selection_drag_active = True
        view._selection_anchor = (120.0, 240.0)
        view._selection_endpoint = (310.0, 240.0)
        view._selection_anchor_page = page
        view._selection_scroll_value = 120.0
        view._on_page_changed = None

        view._on_scroll_changed(_Adjustment(220.0))

        self.assertEqual(
            page._text_selection,
            TextSelection(left=120.0, top=236.0, right=310.0, bottom=344.0),
        )

    def test_active_drag_forwards_wheel_events_from_the_rendered_page(self) -> None:
        # A held pointer grab targets the PDF drawing area, so the page must
        # forward its wheel event to the document's active-selection handler.
        handler = getattr(PdfPage, "_on_scroll", None)
        self.assertIsNotNone(handler)
        page = PdfPage.__new__(PdfPage)
        page._selection_scroll_handler = Mock(return_value=True)
        event = SimpleNamespace(direction=Gdk.ScrollDirection.DOWN, state=0)

        self.assertTrue(handler(page, page, event))

        page._selection_scroll_handler.assert_called_once_with(page, event)

    def test_page_wheel_scrolls_an_active_drag(self) -> None:
        view = PdfDocumentView.__new__(PdfDocumentView)
        view._selection_drag_active = True
        view.scroll = Mock()
        event = SimpleNamespace(direction=Gdk.ScrollDirection.DOWN, state=0)

        self.assertTrue(
            view._on_page_selection_scroll_event(SimpleNamespace(), event)
        )

        view.scroll.assert_called_once_with("line-down")

    def test_cross_page_drag_renders_and_copies_every_intersected_page(self) -> None:
        first_page, first_poppler_page = self._page(
            top=0.0, selected_text="first page"
        )
        second_page, second_poppler_page = self._page(
            top=820.0, selected_text="second page"
        )
        first_page._copy_selection = Mock()

        view = PdfDocumentView.__new__(PdfDocumentView)
        view._pages = [first_page, second_page]
        view._selection_anchor = (120.0, 180.0)
        view._selection_endpoint = (310.0, 1_000.0)
        view._selection_anchor_page = first_page

        view._apply_document_selection()

        self.assertEqual(
            first_page._text_selection,
            TextSelection(left=120.0, top=176.0, right=595.0, bottom=792.0),
        )
        self.assertEqual(
            second_page._text_selection,
            TextSelection(left=0.0, top=0.0, right=310.0, bottom=184.0),
        )
        first_poppler_page.get_selected_text.assert_called_once()
        second_poppler_page.get_selected_text.assert_called_once()
        first_page._copy_selection.assert_called_once_with("first page\nsecond page")
