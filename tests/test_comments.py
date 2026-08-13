from __future__ import annotations

from pathlib import Path
import tempfile
import unittest
import zipfile
from types import SimpleNamespace
from unittest import mock

import see_docx.viewer as viewer


_DOCUMENT_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Read </w:t></w:r>
      <w:commentRangeStart w:id="7"/>
      <w:r><w:t>the results</w:t></w:r>
      <w:commentRangeEnd w:id="7"/>
      <w:r><w:t> carefully.</w:t></w:r>
      <w:r><w:rPr><w:commentReference w:id="7"/></w:rPr></w:r>
    </w:p>
  </w:body>
</w:document>
'''

_COMMENTS_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="7" w:author="Dr. Rivera" w:initials="DR" w:date="2026-08-07T10:30:00Z">
    <w:p><w:r><w:t>Could we make this claim more specific?</w:t></w:r></w:p>
  </w:comment>
</w:comments>
'''

_TABLE_DOCUMENT_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr><w:tc><w:p>
        <w:commentRangeStart w:id="8"/>
        <w:r><w:t>Cell annotation</w:t></w:r>
        <w:commentRangeEnd w:id="8"/>
        <w:r><w:rPr><w:commentReference w:id="8"/></w:rPr></w:r>
      </w:p></w:tc></w:tr>
    </w:tbl>
  </w:body>
</w:document>
'''

_TABLE_COMMENTS_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:comment w:id="8" w:author="Editor"><w:p><w:r><w:t>Table note</w:t></w:r></w:p></w:comment>
</w:comments>
'''

_THREADED_DOCUMENT_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:commentRangeStart w:id="7"/>
      <w:r><w:t>Root target</w:t></w:r>
      <w:commentRangeEnd w:id="7"/>
      <w:r><w:rPr><w:commentReference w:id="7"/></w:rPr></w:r>
    </w:p>
    <w:p>
      <w:commentRangeStart w:id="8"/>
      <w:r><w:t>Reply target</w:t></w:r>
      <w:commentRangeEnd w:id="8"/>
      <w:r><w:rPr><w:commentReference w:id="8"/></w:rPr></w:r>
    </w:p>
  </w:body>
</w:document>
'''

_THREADED_COMMENTS_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:comments xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml">
  <w:comment w:id="7" w:author="Dr. Rivera" w:initials="DR">
    <w:p w14:paraId="ROOT"><w:r><w:t>Root question</w:t></w:r></w:p>
  </w:comment>
  <w:comment w:id="8" w:author="Editor" w:initials="ED">
    <w:p w14:paraId="REPLY"><w:r><w:t>Here is the answer.</w:t></w:r></w:p>
  </w:comment>
</w:comments>
'''

_THREADED_COMMENTS_EXTENDED_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w15:commentsEx
    xmlns:w14="http://schemas.microsoft.com/office/word/2010/wordml"
    xmlns:w15="http://schemas.microsoft.com/office/word/2012/wordml">
  <w15:commentEx w15:paraId="ROOT"/>
  <w15:commentEx w15:paraId="REPLY" w15:paraIdParent="ROOT"/>
</w15:commentsEx>
'''


class CommentSourceTests(unittest.TestCase):
    def test_comment_metadata_and_source_range_are_read_from_ooxml(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "commented.docx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("word/document.xml", _DOCUMENT_XML)
                archive.writestr("word/comments.xml", _COMMENTS_XML)
            source = viewer._docx_rich_text_source(path)

        self.assertIsNotNone(source)
        self.assertEqual(len(source.comments), 1)
        comment = source.comments[0]
        self.assertEqual(comment.author, "Dr. Rivera")
        self.assertEqual(comment.initials, "DR")
        self.assertEqual(comment.text, "Could we make this claim more specific?")
        self.assertEqual(
            source._selection_text[comment.source_start : comment.source_end],
            "the results",
        )

    def test_documents_without_comments_keep_an_empty_comment_model(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "plain.docx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("word/document.xml", _DOCUMENT_XML.replace(
                    b'<w:commentRangeStart w:id="7"/>', b""
                ).replace(
                    b'<w:commentRangeEnd w:id="7"/>', b""
                ))
            source = viewer._docx_rich_text_source(path)

        self.assertIsNotNone(source)
        self.assertEqual(source.comments, ())

    def test_comments_inside_table_cells_keep_their_cell_source_offset(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "table-comment.docx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("word/document.xml", _TABLE_DOCUMENT_XML)
                archive.writestr("word/comments.xml", _TABLE_COMMENTS_XML)
            source = viewer._docx_rich_text_source(path)

        self.assertIsNotNone(source)
        self.assertEqual(len(source.comments), 1)
        comment = source.comments[0]
        self.assertEqual(
            source._selection_text[comment.source_start : comment.source_end],
            "Cell annotation",
        )

    def test_comments_extended_links_replies_into_one_thread(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "threaded.docx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("word/document.xml", _THREADED_DOCUMENT_XML)
                archive.writestr("word/comments.xml", _THREADED_COMMENTS_XML)
                archive.writestr(
                    "word/commentsExtended.xml",
                    _THREADED_COMMENTS_EXTENDED_XML,
                )
            source = viewer._docx_rich_text_source(path)

        self.assertIsNotNone(source)
        self.assertEqual(source.comments[1].parent_comment_id, "7")
        threads = viewer._comment_threads(source.comments)
        self.assertEqual(len(threads), 1)
        self.assertEqual(threads[0].thread_id, "7")
        self.assertEqual([comment.comment_id for comment in threads[0].comments], [
            "7",
            "8",
        ])
        self.assertEqual(threads[0].reply_count, 1)


class CommentFloatGeometryTests(unittest.TestCase):
    def test_page_adjacent_bubble_uses_the_gutter_and_caps_width(self) -> None:
        self.assertEqual(
            viewer._comment_float_geometry(400, 1000, 800, 400, 200),
            (418, 300, 320),
        )

    def test_page_adjacent_bubble_clamps_to_the_canvas_edges(self) -> None:
        self.assertEqual(
            viewer._comment_float_geometry(400, 1000, 800, 10, 400),
            (418, 18, 320),
        )
        self.assertEqual(
            viewer._comment_float_geometry(400, 1000, 800, 795, 400),
            (418, 382, 320),
        )

    def test_page_adjacent_bubble_falls_back_when_the_gutter_is_too_narrow(self) -> None:
        self.assertIsNone(
            viewer._comment_float_geometry(700, 950, 800, 400, 200)
        )

    def test_page_adjacent_bubble_can_cross_the_document_rail_seam(self) -> None:
        self.assertEqual(
            viewer._comment_float_geometry(700, 1080, 800, 400, 200),
            (718, 300, 320),
        )

    def test_page_adjacent_bubble_rejects_offscreen_anchors(self) -> None:
        self.assertIsNone(
            viewer._comment_float_geometry(-1, 1000, 800, 400, 200)
        )
        self.assertIsNone(
            viewer._comment_float_geometry(400, 1000, 800, 801, 200)
        )


class CommentFloatContextTests(unittest.TestCase):
    def test_unanchored_comment_uses_the_nearest_visible_page(self) -> None:
        window = viewer.DocxWindow.__new__(viewer.DocxWindow)
        window._active_comment_layer = SimpleNamespace(
            get_allocated_height=lambda: 800
        )
        pages = [
            SimpleNamespace(
                top=0,
                get_allocated_width=lambda: 600,
                get_allocated_height=lambda: 800,
            ),
            SimpleNamespace(
                top=820,
                get_allocated_width=lambda: 600,
                get_allocated_height=lambda: 800,
            ),
        ]
        window.document = SimpleNamespace(_pages=pages)
        window._translated_point = lambda page, _target, x, y: (
            x,
            page.top + y,
        )

        page_right, context = viewer.DocxWindow._visible_page_float_context(window)

        self.assertEqual(page_right, (600, 400))
        self.assertEqual(context, (600, 400))


class _GhostStyle:
    def __init__(self) -> None:
        self.classes: set[str] = set()

    def add_class(self, name: str) -> None:
        self.classes.add(name)

    def remove_class(self, name: str) -> None:
        self.classes.discard(name)


class _GhostWidget:
    def __init__(self) -> None:
        self.opacity = 1.0

    def set_opacity(self, value: float) -> None:
        self.opacity = value


class _GhostCard:
    def __init__(self) -> None:
        self.style = _GhostStyle()
        self.children = [_GhostWidget(), _GhostWidget()]
        self.size_request: tuple[int, int] | None = None

    def get_style_context(self) -> _GhostStyle:
        return self.style

    def set_size_request(self, width: int, height: int) -> None:
        self.size_request = width, height

    def get_children(self) -> list[_GhostWidget]:
        return self.children


class CommentGhostTests(unittest.TestCase):
    def test_ghosting_keeps_the_original_card_and_exact_dimensions(self) -> None:
        window = viewer.DocxWindow.__new__(viewer.DocxWindow)
        card = _GhostCard()

        viewer.DocxWindow._set_comment_card_ghost(window, card, True, 280, 124)

        self.assertEqual(card.size_request, (280, 124))
        self.assertIn("comment-rail-ghost", card.style.classes)
        self.assertEqual([child.opacity for child in card.children], [0.0, 0.0])

        viewer.DocxWindow._set_comment_card_ghost(window, card, False)

        self.assertEqual(card.size_request, (-1, -1))
        self.assertNotIn("comment-rail-ghost", card.style.classes)
        self.assertEqual([child.opacity for child in card.children], [1.0, 1.0])

    def test_enter_focuses_a_popped_out_comment_without_returning_it_to_the_rail(
        self,
    ) -> None:
        """The visible popout must remain the interaction target after Enter."""

        window = viewer.DocxWindow.__new__(viewer.DocxWindow)
        rail_body = mock.Mock()
        float_body = mock.Mock()
        float_body.is_focus.return_value = True
        float_card = SimpleNamespace(
            get_style_context=lambda: _GhostStyle(),
            grab_focus=mock.Mock(),
        )
        window._comments_focused = True
        window._comments_visible = True
        window._comment_body_focused = False
        window._active_comment_id = "thread-1"
        window._active_comment_float_id = "thread-1"
        window._active_comment_float_card = float_card
        window._active_comment_float_body = float_body
        window._comment_body_scrollers = {"thread-1": rail_body}
        window._comment_annotations = ()
        window._comments_panel = SimpleNamespace(
            get_style_context=lambda: _GhostStyle()
        )
        window._comment_cards = {}
        window._layout_comments = lambda: None
        window._restore_active_comment_card = mock.Mock()

        viewer.DocxWindow._enter_comment_body(window)

        self.assertTrue(window._comment_body_focused)
        window._restore_active_comment_card.assert_not_called()
        float_body.grab_focus.assert_called_once_with()
        rail_body.grab_focus.assert_not_called()
        self.assertIs(viewer.DocxWindow._active_comment_body(window), float_body)

        viewer.DocxWindow._leave_comment_body(window)

        window._restore_active_comment_card.assert_not_called()
        float_card.grab_focus.assert_called_once_with()

    def test_no_space_keeps_the_rail_card_normal_and_does_not_build_a_popout(self) -> None:
        window = viewer.DocxWindow.__new__(viewer.DocxWindow)
        card = _GhostCard()
        card.get_allocated_width = lambda: 280
        card.get_allocated_height = lambda: 124
        window._comments_focused = True
        window._comments_visible = True
        window._comment_body_focused = False
        window._active_comment_id = "thread-1"
        window._active_comment_float_id = None
        window._comment_cards = {"thread-1": card}
        window._comment_float_geometry_for_thread = lambda *_args: None
        window._build_comment_card = lambda *_args, **_kwargs: self.fail(
            "a popout must not be built when there is no room"
        )

        viewer.DocxWindow._promote_active_comment_card(window)

        self.assertNotIn("comment-rail-ghost", card.style.classes)
        self.assertEqual([child.opacity for child in card.children], [1.0, 1.0])


class CommentAnchorMatchTests(unittest.TestCase):
    def test_context_disambiguates_a_short_repeated_comment_range(self) -> None:
        source = "First three words. Second three words."
        start = source.rfind("three")
        end = start + len("three")
        rendered = "First three words.\nSecond three words."

        indices = viewer._comment_rendered_indices(source, start, end, rendered)

        self.assertEqual(
            " ".join("".join(rendered[index] for index in indices).split()),
            "three",
        )

    def test_match_follows_normalized_line_wrapping(self) -> None:
        source = "Before attached phrase after"
        start = source.index("attached")
        end = start + len("attached phrase")
        rendered = "Before attached\nphrase after"

        indices = viewer._comment_rendered_indices(source, start, end, rendered)

        self.assertEqual(
            " ".join("".join(rendered[index] for index in indices).split()),
            "attached phrase",
        )


class CommentFitTests(unittest.TestCase):
    def _window(self) -> tuple[object, SimpleNamespace, list[str]]:
        window = viewer.DocxWindow.__new__(viewer.DocxWindow)
        calls: list[str] = []
        document = SimpleNamespace(
            has_document=True,
            zoom=1.25,
            widget=SimpleNamespace(get_vadjustment=lambda: object()),
        )

        def fit_to_width() -> None:
            calls.append("fit")
            document.zoom = 0.82

        def set_zoom(zoom: float) -> None:
            calls.append(f"restore:{zoom}")
            document.zoom = zoom

        document.fit_to_width = fit_to_width
        document.set_zoom = set_zoom
        window.document = document
        window.get_visible = lambda: True
        window._comments_visible = True
        window._comments_available = True
        window._comments_zoom_before_fit = None
        window._comments_auto_fit_zoom = None
        window._schedule_active_comment_position = lambda: None
        return window, document, calls

    def test_comment_rail_fits_once_and_restores_automatic_zoom(self) -> None:
        window, document, calls = self._window()

        viewer.DocxWindow._fit_document_for_comments(window)
        self.assertEqual(document.zoom, 0.82)
        self.assertEqual(window._comments_zoom_before_fit, 1.25)
        self.assertEqual(window._comments_auto_fit_zoom, 0.82)

        viewer.DocxWindow._fit_document_for_comments(window)
        # A second layout pass is allowed to recompute the target if the
        # document column changed size while the rail was revealing.
        self.assertEqual(calls, ["fit", "fit"])

        viewer.DocxWindow._restore_document_zoom_after_comments(window)
        self.assertEqual(document.zoom, 1.25)
        self.assertEqual(calls, ["fit", "fit", "restore:1.25"])

    def test_manual_zoom_change_is_preserved_when_comments_close(self) -> None:
        window, document, calls = self._window()

        viewer.DocxWindow._fit_document_for_comments(window)
        document.zoom = 1.0
        viewer.DocxWindow._restore_document_zoom_after_comments(window)

        self.assertEqual(document.zoom, 1.0)
        self.assertEqual(calls, ["fit"])

    def test_queued_fit_stops_when_the_window_is_hidden(self) -> None:
        window, _document, calls = self._window()
        window.get_visible = lambda: False

        viewer.DocxWindow._fit_document_for_comments(window)

        self.assertEqual(calls, [])


class _CommentMarkContext:
    def __init__(self) -> None:
        self.colors: list[tuple[float, float, float, float]] = []

    def save(self) -> None:
        pass

    def restore(self) -> None:
        pass

    def set_operator(self, _operator: object) -> None:
        pass

    def set_source_rgba(self, *color: float) -> None:
        self.colors.append(color)

    def rectangle(self, *_args: float) -> None:
        pass

    def fill(self) -> None:
        pass

    def set_line_width(self, _width: float) -> None:
        pass

    def move_to(self, *_args: float) -> None:
        pass

    def line_to(self, *_args: float) -> None:
        pass

    def stroke(self) -> None:
        pass


class CommentMarkRenderTests(unittest.TestCase):
    def test_active_comment_gets_high_contrast_paper_treatment(self) -> None:
        page = viewer.PdfPage.__new__(viewer.PdfPage)
        page._comment_marks = (
            viewer.CommentMark("active", (0.0, 0.0, 10.0, 10.0)),
            viewer.CommentMark("other", (20.0, 0.0, 30.0, 10.0)),
            viewer.CommentMark("other-duplicate", (20.0, 0.0, 30.0, 10.0)),
            viewer.CommentMark("stacked-inactive", (0.0, 0.0, 10.0, 10.0)),
        )
        page._active_comment_id = "active"
        page._zoom = 1.0
        context = _CommentMarkContext()
        with mock.patch.object(
            viewer,
            "_theme_palette",
            return_value={"accent": "#00ff00", "highlight": "#ffff00"},
        ):
            viewer.PdfPage._draw_comment_marks(page, context)

        self.assertEqual(len(context.colors), 6)
        self.assertEqual(
            [color[:3] for color in context.colors],
            [(0.0, 1.0, 0.0)] * 6,
        )
        self.assertEqual(
            [round(color[3], 2) for color in context.colors],
            [0.03, 0.18, 0.03, 0.18, 0.34, 0.99],
        )


class _VisiblePanel:
    def get_visible(self) -> bool:
        return False


class _Revealer:
    def set_reveal_child(self, value: bool) -> None:
        self.revealed = value


class _Layer:
    def set_visible(self, value: bool) -> None:
        self.visible = value


class _CommentToggleWindow:
    _on_key_press = viewer.DocxWindow._on_key_press
    _toggle_comments = viewer.DocxWindow._toggle_comments

    def __init__(self) -> None:
        self._search_panel = _VisiblePanel()
        self._page_jump_panel = _VisiblePanel()
        self._outline_panel = _VisiblePanel()
        self._export_panel = _VisiblePanel()
        self._search_entry = SimpleNamespace(is_focus=lambda: False)
        self._page_jump_entry = SimpleNamespace(is_focus=lambda: False)
        self._url_hint_targets: dict[str, str] = {}
        self._pending_g = False
        self._comments_visible = True
        self._comments_focused = False
        self._comment_body_focused = False
        self._comments_revealer = _Revealer()
        self._comment_line_layer = _Layer()
        self._layout_comments = lambda: None
        self._apply_comment_sizing = lambda: None
        self._comment_annotations = (SimpleNamespace(comment_id="one"),)
        self._active_comment_index = 0
        self._active_comment_id = "one"
        self._move_comment_calls: list[int] = []
        self._comment_list_scroll_calls: list[int] = []
        self._comment_body_scroll_calls: list[int] = []
        self._activate_comment_calls: list[tuple[int, bool]] = []
        self._scroll_active_comment_into_view = lambda: False
        self._fit_document_for_comments = lambda: False
        self._restore_document_zoom_after_comments = lambda: None
        self._active_comment_body = lambda: (
            SimpleNamespace() if self._comment_body_focused else None
        )
        self.document = SimpleNamespace(widget=SimpleNamespace(grab_focus=lambda: None))

    def _move_comment_selection(self, direction: int) -> None:
        self._move_comment_calls.append(direction)

    def _focus_comments(self) -> None:
        self._comments_focused = True
        self._comment_body_focused = False

    def _scroll_comment_list(self, direction: int) -> None:
        self._comment_list_scroll_calls.append(direction)

    def _scroll_comment_body(self, _body: object, direction: int) -> None:
        self._comment_body_scroll_calls.append(direction)

    def _enter_comment_body(self) -> None:
        self._comment_body_focused = True

    def _activate_comment(self, index: int, *, reveal_document: bool) -> None:
        self._activate_comment_calls.append((index, reveal_document))

    def _leave_comment_body(self) -> None:
        self._comment_body_focused = False

    def _blur_comments(self) -> None:
        self._comments_focused = False
        self._comment_body_focused = False


class CommentToggleTests(unittest.TestCase):
    def test_comment_focus_levels_unwind_with_escape(self) -> None:
        window = _CommentToggleWindow()

        self.assertTrue(
            window._on_key_press(window, SimpleNamespace(keyval=viewer.Gdk.KEY_c, state=0))
        )
        self.assertTrue(window._comments_focused)
        self.assertFalse(window._comment_body_focused)
        self.assertTrue(
            window._on_key_press(window, SimpleNamespace(keyval=viewer.Gdk.KEY_c, state=0))
        )
        self.assertEqual(window._move_comment_calls, [])
        self.assertTrue(
            window._on_key_press(window, SimpleNamespace(keyval=viewer.Gdk.KEY_j, state=0))
        )
        self.assertTrue(
            window._on_key_press(window, SimpleNamespace(keyval=viewer.Gdk.KEY_k, state=0))
        )
        self.assertEqual(window._move_comment_calls, [1, -1])
        self.assertTrue(
            window._on_key_press(
                window,
                SimpleNamespace(
                    keyval=viewer.Gdk.KEY_d,
                    state=viewer.Gdk.ModifierType.CONTROL_MASK,
                ),
            )
        )
        self.assertEqual(window._comment_list_scroll_calls, [1])
        self.assertTrue(
            window._on_key_press(
                window, SimpleNamespace(keyval=viewer.Gdk.KEY_Return, state=0)
            )
        )
        self.assertTrue(window._comment_body_focused)
        self.assertTrue(
            window._on_key_press(window, SimpleNamespace(keyval=viewer.Gdk.KEY_j, state=0))
        )
        self.assertEqual(window._comment_body_scroll_calls, [1])

        self.assertTrue(
            window._on_key_press(window, SimpleNamespace(keyval=viewer.Gdk.KEY_Escape, state=0))
        )
        self.assertTrue(window._comments_focused)
        self.assertFalse(window._comment_body_focused)
        self.assertTrue(
            window._on_key_press(window, SimpleNamespace(keyval=viewer.Gdk.KEY_Escape, state=0))
        )
        self.assertFalse(window._comments_focused)

    def test_v_toggles_comments_off_then_back_on(self) -> None:
        window = _CommentToggleWindow()

        self.assertTrue(
            window._on_key_press(window, SimpleNamespace(keyval=viewer.Gdk.KEY_v, state=0))
        )
        self.assertFalse(window._comments_visible)
        self.assertFalse(window._comment_line_layer.visible)
        self.assertFalse(window._comments_revealer.revealed)

        self.assertTrue(
            window._on_key_press(window, SimpleNamespace(keyval=viewer.Gdk.KEY_v, state=0))
        )
        self.assertTrue(window._comments_visible)
        self.assertTrue(window._comment_line_layer.visible)
        self.assertTrue(window._comments_revealer.revealed)

    def test_v_does_not_open_an_unavailable_comment_rail(self) -> None:
        window = _CommentToggleWindow()
        window._comment_annotations = ()
        window._comments_available = False
        window._comments_visible = False
        window._comments_revealer.revealed = False
        window._comment_line_layer.visible = False

        self.assertTrue(
            window._on_key_press(
                window, SimpleNamespace(keyval=viewer.Gdk.KEY_v, state=0)
            )
        )
        self.assertFalse(window._comments_visible)
        self.assertFalse(window._comment_line_layer.visible)
        self.assertFalse(window._comments_revealer.revealed)

    def test_gg_and_G_jump_to_the_first_and_last_comment_thread(self) -> None:
        window = _CommentToggleWindow()
        window._comments_focused = True
        window._comment_annotations = (
            SimpleNamespace(comment_id="one"),
            SimpleNamespace(comment_id="two"),
        )

        self.assertTrue(
            window._on_key_press(
                window, SimpleNamespace(keyval=viewer.Gdk.KEY_g, state=0)
            )
        )
        self.assertEqual(window._activate_comment_calls, [])
        self.assertTrue(
            window._on_key_press(
                window, SimpleNamespace(keyval=viewer.Gdk.KEY_g, state=0)
            )
        )
        self.assertEqual(window._activate_comment_calls, [(0, True)])
        self.assertTrue(
            window._on_key_press(
                window, SimpleNamespace(keyval=viewer.Gdk.KEY_G, state=0)
            )
        )
        self.assertEqual(window._activate_comment_calls, [(0, True), (1, True)])

    def test_comment_selection_stops_at_both_ends(self) -> None:
        window = SimpleNamespace(
            _comment_annotations=(object(), object(), object()),
            _comments_visible=True,
            _comments_focused=True,
            _comment_body_focused=False,
            _active_comment_index=2,
            _activate_comment=mock.Mock(),
        )

        viewer.DocxWindow._move_comment_selection(window, 1)
        window._activate_comment.assert_called_once_with(
            2,
            reveal_document=True,
        )
        window._activate_comment.reset_mock()
        window._active_comment_index = 0
        viewer.DocxWindow._move_comment_selection(window, -1)
        window._activate_comment.assert_called_once_with(
            0,
            reveal_document=True,
        )


if __name__ == "__main__":
    unittest.main()
