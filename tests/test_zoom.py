from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import Mock, patch
import unittest

from see_docx.viewer import DocxWindow, Gdk, MAX_ZOOM, PdfDocumentView, ZOOM_STEP


class _Page:
    def __init__(self) -> None:
        self.zooms: list[float] = []

    def set_zoom(self, zoom: float) -> None:
        self.zooms.append(zoom)


class ZoomTests(unittest.TestCase):
    def test_window_resize_scales_current_zoom_in_both_directions(self) -> None:
        """Page and surrounding whitespace must retain their width ratio."""

        handler = getattr(DocxWindow, "_on_window_size_allocate", None)
        self.assertIsNotNone(handler, "window resize has no adaptive zoom policy")
        document = SimpleNamespace(
            has_document=True,
            zoom=1.25,
        )

        def set_zoom(zoom: float, **_kwargs: object) -> bool:
            document.zoom = zoom
            return True

        document.set_zoom = set_zoom
        window = DocxWindow.__new__(DocxWindow)
        window.document = document
        window._viewport_resize_width = 933
        window._comments_zoom_before_fit = None
        window._comments_auto_fit_zoom = None
        window._outline_zoom_before_open = None

        handler(window, window, SimpleNamespace(width=1920))
        self.assertAlmostEqual(document.zoom, 1.25 * 1920 / 933)

        handler(window, window, SimpleNamespace(width=933))
        self.assertAlmostEqual(document.zoom, 1.25)

    def test_resizes_existing_pages_instead_of_rebuilding_the_document(self) -> None:
        view = object.__new__(PdfDocumentView)
        view._document = object()
        view.zoom = 1.0
        view._pages = [_Page(), _Page()]
        view._pages_box = SimpleNamespace(queue_resize=lambda: None)
        view.capture_position = lambda: object()
        view.set_document = lambda _document: self.fail("zoom must not rebuild pages")

        with patch("see_docx.viewer.GLib.idle_add") as idle_add:
            self.assertTrue(view.set_zoom(1.1))

        self.assertEqual([page.zooms for page in view._pages], [[1.1], [1.1]])
        idle_add.assert_called_once()

    def test_manual_zoom_keeps_its_limit_while_adaptive_zoom_can_exceed_it(self) -> None:
        view = object.__new__(PdfDocumentView)
        view._document = object()
        view.zoom = 1.0
        view._pages = [_Page()]
        view._pages_box = SimpleNamespace(queue_resize=lambda: None)
        view.capture_position = lambda: object()

        with patch("see_docx.viewer.GLib.idle_add"):
            self.assertTrue(view.set_zoom(99.0))
            self.assertEqual(view.zoom, MAX_ZOOM)
            self.assertTrue(view.set_zoom(2.5, maximum=None))
            self.assertEqual(view.zoom, 2.5)

    def test_resize_scales_sidebar_zoom_restore_points(self) -> None:
        document = SimpleNamespace(has_document=True, zoom=0.8)

        def set_zoom(zoom: float, **_kwargs: object) -> bool:
            document.zoom = zoom
            return True

        document.set_zoom = set_zoom
        window = DocxWindow.__new__(DocxWindow)
        window.document = document
        window._viewport_resize_width = 800
        window._comments_zoom_before_fit = 1.0
        window._comments_auto_fit_zoom = 0.8
        window._outline_zoom_before_open = 1.1

        DocxWindow._on_window_size_allocate(
            window,
            window,
            SimpleNamespace(width=1200),
        )

        self.assertAlmostEqual(document.zoom, 1.2)
        self.assertAlmostEqual(window._comments_auto_fit_zoom, 1.2)
        self.assertAlmostEqual(window._comments_zoom_before_fit, 1.5)
        self.assertAlmostEqual(window._outline_zoom_before_open, 1.65)

    def test_control_mouse_wheel_zooms_the_document_without_scrolling(self) -> None:
        handler = getattr(PdfDocumentView, "_on_scroll_event", None)
        self.assertIsNotNone(handler)
        view = SimpleNamespace(zoom=1.25, set_zoom=Mock(return_value=True))
        event = SimpleNamespace(
            state=Gdk.ModifierType.CONTROL_MASK,
            direction=Gdk.ScrollDirection.UP,
        )

        self.assertTrue(handler(view, view, event))
        view.set_zoom.assert_called_once_with(1.25 + ZOOM_STEP)

    def test_mouse_wheel_without_control_keeps_its_normal_scroll_behavior(self) -> None:
        handler = getattr(PdfDocumentView, "_on_scroll_event", None)
        self.assertIsNotNone(handler)
        view = SimpleNamespace(zoom=1.25, set_zoom=Mock(return_value=True))
        event = SimpleNamespace(state=0, direction=Gdk.ScrollDirection.UP)

        self.assertFalse(handler(view, view, event))
        view.set_zoom.assert_not_called()
