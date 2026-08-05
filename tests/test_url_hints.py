from __future__ import annotations

from types import SimpleNamespace
import unittest
from unittest.mock import patch

from see_docx.viewer import (
    DocxWindow,
    Gdk,
    HINT_CHARS,
    PdfPage,
    Poppler,
    UrlLink,
    _page_url_links,
    hint_codes,
)


class _HintLabel:
    def __init__(self) -> None:
        self.visible = True

    def set_visible(self, visible: bool) -> None:
        self.visible = visible


class _HintLayer:
    def __init__(self) -> None:
        self.hidden = False

    def hide(self) -> None:
        self.hidden = True


class _Panel:
    def get_visible(self) -> bool:
        return False


class _Entry:
    def is_focus(self) -> bool:
        return False


class _UrlHintWindow:
    """The real URL-hint handlers with GTK display objects replaced."""

    _on_key_press = DocxWindow._on_key_press
    _filter_url_hints = DocxWindow._filter_url_hints
    _erase_url_hint_character = DocxWindow._erase_url_hint_character

    def __init__(self) -> None:
        self._search_panel = _Panel()
        self._page_jump_panel = _Panel()
        self._outline_panel = _Panel()
        self._search_entry = _Entry()
        self._page_jump_entry = _Entry()
        self._url_hint_targets: dict[str, str] = {}
        self._url_hint_labels: dict[str, _HintLabel] = {}
        self._url_hint_prefix = ""
        self._pending_g = False
        self.shown = False
        self.opened: list[str] = []

    def _show_url_hints(self) -> None:
        self.shown = True

    def _hide_url_hints(self) -> None:
        self._url_hint_targets.clear()
        self._url_hint_labels.clear()
        self._url_hint_prefix = ""

    def _open_url(self, uri: str) -> None:
        self.opened.append(uri)

    def _set_url_hint_text(self, _label: _HintLabel, _code: str) -> None:
        pass


class UrlHintTests(unittest.TestCase):
    def test_keeps_only_external_uri_actions_and_normalizes_their_rectangles(self) -> None:
        uri_action = SimpleNamespace(
            type=Poppler.ActionType.URI,
            uri=SimpleNamespace(uri="https://example.com/research"),
        )
        ignored_action = SimpleNamespace(
            type=Poppler.ActionType.GOTO_DEST,
            uri=SimpleNamespace(uri="https://example.com/ignored"),
        )
        page = SimpleNamespace(
            get_link_mapping=lambda: [
                SimpleNamespace(
                    action=uri_action,
                    area=SimpleNamespace(x1=240, y1=620, x2=80, y2=640),
                ),
                SimpleNamespace(
                    action=ignored_action,
                    area=SimpleNamespace(x1=1, y1=2, x2=3, y2=4),
                ),
            ]
        )

        self.assertEqual(
            _page_url_links(page),
            [
                UrlLink(
                    uri="https://example.com/research",
                    left=80.0,
                    bottom=620.0,
                    right=240.0,
                    top=640.0,
                )
            ],
        )

    def test_hint_codes_match_see_mail_and_remain_prefix_free(self) -> None:
        codes = hint_codes(10)

        self.assertEqual(codes, ["a", "ls", "s", "d", "f", "g", "h", "j", "k", "la"])
        self.assertTrue(all(set(code) <= set(HINT_CHARS) for code in codes))
        self.assertFalse(
            any(
                code != other and other.startswith(code)
                for code in codes
                for other in codes
            )
        )

    def test_places_a_url_hint_at_the_pdf_links_top_edge(self) -> None:
        page = PdfPage.__new__(PdfPage)
        page._height = 792.0
        page._zoom = 1.25
        link = UrlLink("https://example.com", 80.0, 620.0, 240.0, 640.0)

        self.assertEqual(page.url_link_position(link), (100, 190))
        self.assertEqual(page.url_link_size(link), (200, 25))

    def test_skips_a_link_when_gtk_cannot_translate_its_layout_coordinates(self) -> None:
        # A freshly rebuilt ScrolledWindow can temporarily have no page-to-
        # overlay transform. Hint mode must wait for the next f, not crash.
        page = SimpleNamespace(
            url_link_position=lambda _link: (100, 190),
            translate_coordinates=lambda _layer, _x, _y: None,
        )
        window = SimpleNamespace(_url_hint_layer=object())

        self.assertIsNone(
            DocxWindow._url_hint_position(
                window,
                page,
                UrlLink("https://example.com", 80.0, 620.0, 240.0, 640.0),
            )
        )

    def test_dismissing_hints_keeps_the_pass_through_layer_allocated(self) -> None:
        # Its coordinates must remain available before the next f shows a
        # label; hiding the overlay collapses it to an unusable 1×1 widget.
        layer = _HintLayer()
        window = SimpleNamespace(
            _url_hint_labels={},
            _url_hint_targets={"a": "https://example.com"},
            _url_hint_prefix="a",
            _url_hint_layer=layer,
        )

        DocxWindow._hide_url_hints(window)

        self.assertFalse(layer.hidden)
        self.assertEqual(window._url_hint_targets, {})
        self.assertEqual(window._url_hint_prefix, "")

    def test_f_opens_hint_mode_and_a_complete_hint_launches_its_uri(self) -> None:
        window = _UrlHintWindow()

        self.assertTrue(
            window._on_key_press(
                window, SimpleNamespace(keyval=Gdk.KEY_f, state=0)
            )
        )
        self.assertTrue(window.shown)

        window._url_hint_targets = {"as": "https://example.com/research"}
        window._url_hint_labels = {"as": _HintLabel()}
        self.assertTrue(
            window._on_key_press(
                window, SimpleNamespace(keyval=Gdk.KEY_a, state=0)
            )
        )
        self.assertEqual(window._url_hint_prefix, "a")
        self.assertTrue(
            window._on_key_press(
                window, SimpleNamespace(keyval=Gdk.KEY_s, state=0)
            )
        )
        self.assertEqual(window.opened, ["https://example.com/research"])
        self.assertEqual(window._url_hint_targets, {})

    def test_launches_the_selected_uri_with_the_default_desktop_handler(self) -> None:
        window = SimpleNamespace(_set_status=lambda _text: None)

        with patch("see_docx.viewer.Gio.AppInfo.launch_default_for_uri") as launch:
            DocxWindow._open_url(window, "mailto:research@example.com")

        launch.assert_called_once_with("mailto:research@example.com", None)
