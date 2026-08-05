from __future__ import annotations

from dataclasses import dataclass
import unittest

from see_docx.viewer import SearchMatch, _document_search


@dataclass
class _Rectangle:
    x1: float
    y1: float
    x2: float
    y2: float


class _Page:
    def __init__(self, matches: dict[str, list[_Rectangle]]) -> None:
        self.matches = matches

    def find_text(self, query: str) -> list[_Rectangle]:
        return self.matches.get(query, [])


class _Document:
    def __init__(self, pages: list[_Page]) -> None:
        self.pages = pages

    def get_n_pages(self) -> int:
        return len(self.pages)

    def get_page(self, index: int) -> _Page:
        return self.pages[index]


class SearchTests(unittest.TestCase):
    def test_returns_every_native_text_match_with_its_page_destination(self) -> None:
        document = _Document(
            [
                _Page({"needle": [_Rectangle(10.0, 696.0, 48.0, 710.0)]}),
                _Page(
                    {
                        "needle": [
                            _Rectangle(20.0, 287.0, 56.0, 301.5),
                            _Rectangle(12.0, 60.0, 50.0, 74.0),
                        ]
                    }
                ),
            ]
        )

        self.assertEqual(
            _document_search(document, "needle"),
            [
                SearchMatch(page_index=0, left=10.0, bottom=696.0, right=48.0, top=710.0),
                SearchMatch(page_index=1, left=20.0, bottom=287.0, right=56.0, top=301.5),
                SearchMatch(page_index=1, left=12.0, bottom=60.0, right=50.0, top=74.0),
            ],
        )
