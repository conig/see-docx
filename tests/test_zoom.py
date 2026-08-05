from __future__ import annotations

from types import SimpleNamespace
from unittest.mock import Mock, patch
import unittest

from see_docx.viewer import Gdk, PdfDocumentView, ZOOM_STEP


class _Page:
    def __init__(self) -> None:
        self.zooms: list[float] = []

    def set_zoom(self, zoom: float) -> None:
        self.zooms.append(zoom)


class ZoomTests(unittest.TestCase):
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
