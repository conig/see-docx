"""Pure position capture and restoration for a re-rendered document."""

from __future__ import annotations

from dataclasses import dataclass


@dataclass(frozen=True)
class PageGeometry:
    """The vertical bounds of a visible rendered page."""

    top: float
    height: float

    @property
    def bottom(self) -> float:
        return self.top + self.height


@dataclass(frozen=True)
class DocumentPosition:
    """A page-relative position with a document-relative fallback."""

    page_index: int | None
    page_fraction: float | None
    document_fraction: float


def _clamp(value: float, lower: float, upper: float) -> float:
    return min(max(value, lower), upper)


def capture_position(
    scroll_value: float,
    pages: list[PageGeometry],
    *,
    maximum_scroll: float,
) -> DocumentPosition:
    """Capture the page and within-page offset currently at the viewport top."""

    maximum_scroll = max(maximum_scroll, 0.0)
    scroll_value = _clamp(scroll_value, 0.0, maximum_scroll)
    document_fraction = scroll_value / maximum_scroll if maximum_scroll else 0.0

    for page_index, page in enumerate(pages):
        if page.height > 0 and page.top <= scroll_value < page.bottom:
            return DocumentPosition(
                page_index=page_index,
                page_fraction=(scroll_value - page.top) / page.height,
                document_fraction=document_fraction,
            )

    return DocumentPosition(
        page_index=None,
        page_fraction=None,
        document_fraction=document_fraction,
    )


def page_index_at_scroll(
    scroll_value: float,
    pages: list[PageGeometry],
    *,
    maximum_scroll: float,
) -> int | None:
    """Return the page represented by a scroll location for page navigation."""

    scroll_value = _clamp(scroll_value, 0.0, max(maximum_scroll, 0.0))
    last_page_index: int | None = None
    for page_index, page in enumerate(pages):
        if page.height <= 0:
            continue
        last_page_index = page_index
        if scroll_value < page.bottom:
            return page_index
    return last_page_index


def restore_position(
    position: DocumentPosition,
    pages: list[PageGeometry],
    *,
    maximum_scroll: float,
) -> float:
    """Restore a captured location, preferring the same page and offset.

    If a re-render removes that page, restore the equivalent whole-document
    percentage.  This keeps an edit near the current location from jumping to
    the document start while still behaving sensibly after a large reflow.
    """

    maximum_scroll = max(maximum_scroll, 0.0)
    if (
        position.page_index is not None
        and position.page_fraction is not None
        and 0 <= position.page_index < len(pages)
    ):
        page = pages[position.page_index]
        target = page.top + _clamp(position.page_fraction, 0.0, 1.0) * page.height
    else:
        target = _clamp(position.document_fraction, 0.0, 1.0) * maximum_scroll
    return _clamp(target, 0.0, maximum_scroll)
