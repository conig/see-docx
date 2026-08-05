from __future__ import annotations

from types import SimpleNamespace
import unittest
from unittest.mock import patch

from see_docx.viewer import (
    OutlineEntry,
    _document_outline,
    _outline_default_expansion_depth,
    _outline_entries_from_iter,
)


class _Index:
    def __init__(self, nodes: list[tuple[object, _Index | None]]) -> None:
        self.nodes = nodes
        self.position = 0

    def get_action(self) -> object:
        return self.nodes[self.position][0]

    def get_child(self) -> _Index | None:
        return self.nodes[self.position][1]

    def next(self) -> bool:
        if self.position + 1 >= len(self.nodes):
            return False
        self.position += 1
        return True


class _Document:
    def __init__(self, page_count: int) -> None:
        self.page_count = page_count

    def get_n_pages(self) -> int:
        return self.page_count

    def find_dest(self, _name: str) -> None:
        return None

    def get_page(self, _index: int) -> object:
        return object()


def _action(
    title: str, page_index: int, top: float, *, change_top: bool = True
) -> object:
    destination = SimpleNamespace(
        type=1,
        named_dest=None,
        page_num=page_index,
        change_top=change_top,
        top=top,
    )
    return SimpleNamespace(
        type=2,
        goto_dest=SimpleNamespace(title=title, dest=destination),
    )


class OutlineTests(unittest.TestCase):
    def test_documents_without_bookmarks_have_an_empty_outline(self) -> None:
        # Current Poppler bindings raise TypeError, rather than returning
        # None, when a PDF has no outline at all. Such documents must still
        # render so their URI links can be hinted and opened.
        with patch("see_docx.viewer.Poppler") as poppler:
            poppler.IndexIter.new.side_effect = TypeError("constructor returned NULL")

            self.assertEqual(_document_outline(object()), [])

    def test_keeps_a_valid_libreoffice_xyz_destination_when_change_top_is_unset(self) -> None:
        # LibreOffice emits XYZ outline destinations with a real Y coordinate
        # but marks ``change_top`` false. Dropping that coordinate skips both
        # contextual navigation and the arrival animation.
        page = SimpleNamespace(get_size=lambda: (595.0, 792.0))
        document = SimpleNamespace(
            get_n_pages=lambda: 1,
            get_page=lambda _index: page,
            find_dest=lambda _name: None,
        )

        entries = _outline_entries_from_iter(
            document,
            _Index([(_action("Methods", 1, 420.0, change_top=False), None)]),
        )

        self.assertEqual(entries[0].top, 420.0)

    def test_outline_destination_ignores_matching_body_text(self) -> None:
        # The title can recur in body text, but outline navigation must remain
        # wholly determined by the PDF bookmark destination.
        page = SimpleNamespace(
            find_text=lambda _title: self.fail("outline navigation must not search body text"),
        )
        document = SimpleNamespace(
            get_n_pages=lambda: 1,
            get_page=lambda _index: page,
            find_dest=lambda _name: None,
        )

        entries = _outline_entries_from_iter(document, _Index([(_action("Methods", 1, 420), None)]))

        self.assertEqual(entries[0].top, 420.0)

    def test_default_expansion_stops_before_ten_visible_headings(self) -> None:
        entries = [
            *(OutlineEntry(f"Chapter {number}", 0, None, 0) for number in range(4)),
            *(OutlineEntry(f"Section {number}", 0, None, 1) for number in range(5)),
            *(OutlineEntry(f"Detail {number}", 0, None, 2) for number in range(3)),
        ]

        self.assertEqual(_outline_default_expansion_depth(entries), 1)
        self.assertEqual(
            _outline_default_expansion_depth(
                [
                    *(OutlineEntry(f"Chapter {number}", 0, None, 0) for number in range(9)),
                    OutlineEntry("Section", 0, None, 1),
                ]
            ),
            0,
        )

    def test_flattens_nested_headings_and_keeps_their_destinations(self) -> None:
        child = _Index([(_action("Details", 1, 620), None)])
        root = _Index(
            [
                (_action("Introduction", 1, 720), child),
                (_action("Conclusion", 3, 510), None),
            ]
        )

        entries = _outline_entries_from_iter(_Document(3), root)

        self.assertEqual(
            [(entry.title, entry.page_index, entry.top, entry.depth) for entry in entries],
            [
                ("Introduction", 0, 720.0, 0),
                ("Details", 0, 620.0, 1),
                ("Conclusion", 2, 510.0, 0),
            ],
        )
