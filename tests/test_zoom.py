from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import Mock, patch
import unittest

from see_docx.viewer import (
    DocxWindow,
    Gdk,
    MAX_ZOOM,
    MIN_FIT_ZOOM,
    PAGE_MARGIN,
    PdfDocumentView,
    ZOOM_STEP,
)


class _Page:
    def __init__(self, width: float = 0.0, height: float = 0.0) -> None:
        self._width = width
        self._height = height
        self.zooms: list[float] = []

    def set_zoom(self, zoom: float) -> None:
        self.zooms.append(zoom)


class _HiddenPanel:
    def get_visible(self) -> bool:
        return False


class _Entry:
    def is_focus(self) -> bool:
        return False


class _WidthFitKeyWindow:
    _on_key_press = DocxWindow._on_key_press

    def __init__(self) -> None:
        self._search_panel = _HiddenPanel()
        self._page_jump_panel = _HiddenPanel()
        self._export_panel = _HiddenPanel()
        self._outline_panel = _HiddenPanel()
        self._search_entry = _Entry()
        self._page_jump_entry = _Entry()
        self._url_hint_targets: dict[str, str] = {}
        self._pending_g = False
        self.document = SimpleNamespace(
            zoom_to_width=Mock(return_value=True),
            zoom_to_height=Mock(return_value=True),
        )


class ZoomTests(unittest.TestCase):
    def test_z_fits_the_page_to_the_current_central_pane_width(self) -> None:
        window = _WidthFitKeyWindow()

        self.assertTrue(
            window._on_key_press(
                window,
                SimpleNamespace(keyval=Gdk.KEY_z, state=0),
            )
        )

        window.document.zoom_to_width.assert_called_once_with()

    def test_uppercase_z_fits_the_complete_page_height(self) -> None:
        window = _WidthFitKeyWindow()

        self.assertTrue(
            window._on_key_press(
                window,
                SimpleNamespace(keyval=Gdk.KEY_Z, state=Gdk.ModifierType.SHIFT_MASK),
            )
        )

        window.document.zoom_to_height.assert_called_once_with()

    def test_width_fit_uses_the_current_central_pane_allocation_exactly(self) -> None:
        view = object.__new__(PdfDocumentView)
        view._document = object()
        view._pages = [_Page(width=600.0)]
        pane = SimpleNamespace(get_allocated_width=Mock(return_value=972))
        view.widget = pane
        view.set_zoom = Mock(return_value=True)

        self.assertTrue(view.zoom_to_width())
        view.set_zoom.assert_called_once_with(
            (972 - 2 * PAGE_MARGIN) / 600.0,
            minimum=MIN_FIT_ZOOM,
            maximum=None,
        )

        pane.get_allocated_width.return_value = 672
        view.zoom_to_width()
        self.assertAlmostEqual(view.set_zoom.call_args.args[0], 1.0)

    def test_height_fit_uses_the_current_document_viewport_exactly(self) -> None:
        view = object.__new__(PdfDocumentView)
        view._document = object()
        view._pages = [_Page(width=600.0, height=800.0)]
        pane = SimpleNamespace(get_allocated_height=Mock(return_value=872))
        view.widget = pane
        view.set_zoom = Mock(return_value=True)

        self.assertTrue(view.zoom_to_height())
        view.set_zoom.assert_called_once_with(
            (872 - 2 * PAGE_MARGIN) / 800.0,
            minimum=MIN_FIT_ZOOM,
            maximum=None,
        )

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
