#!/usr/bin/env python3
"""Exercise a real watched DOCX refresh and assert the reading location holds.

This is deliberately a manual smoke test: it needs a GTK display, Poppler, and
LibreOffice. It copies the supplied document to a private temporary file,
touches only that copy, and exits non-zero if the refresh returns to the top.
"""

from __future__ import annotations

import argparse
from pathlib import Path
import shutil
import tempfile
from types import SimpleNamespace

import gi

gi.require_version("Gdk", "3.0")
gi.require_version("Gtk", "3.0")
gi.require_version("Pango", "1.0")

from gi.repository import Gdk, Gio, GLib, Gtk, Pango

from see_docx.viewer import (
    APPLICATION_ID,
    OUTLINE_NAV_SPACING,
    OUTLINE_LOCATOR_ARRIVAL_MS,
    PAGE_GAP,
    DocxWindow,
    OutlineEntry,
    _compact_path,
)


TIMEOUT_MS = 20_000


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("path", type=Path, help="A multi-page DOCX fixture")
    parser.add_argument(
        "--outline-animation-only",
        action="store_true",
        help="Verify outline arrival frames without exercising live refresh",
    )
    parser.add_argument(
        "--scroll-navigation-only",
        action="store_true",
        help="Verify initial progress plus keyboard and outline navigation",
    )
    parser.add_argument(
        "--outline-count-only",
        action="store_true",
        help="Verify count-prefixed outline j/k navigation",
    )
    parser.add_argument(
        "--search-cancel-only",
        action="store_true",
        help="Verify Escape clears a committed document search session",
    )
    arguments = parser.parse_args()
    source = arguments.path.resolve()
    if not source.is_file():
        parser.error(f"No such DOCX: {source}")

    with tempfile.TemporaryDirectory(prefix="see-docx-ui-smoke-") as directory:
        watched = Path(directory) / "watched.docx"
        shutil.copy2(source, watched)
        result: dict[str, str] = {}
        state: dict[str, object] = {
            "replaced": False,
            "expected": None,
            "initial_revision": None,
            "before": None,
        }
        application = Gtk.Application(
            application_id=f"{APPLICATION_ID}.smoke",
            flags=Gio.ApplicationFlags.NON_UNIQUE,
        )

        def finish(window: DocxWindow, message: str | None = None) -> bool:
            if message:
                result["error"] = message
            window.close()
            application.quit()
            return GLib.SOURCE_REMOVE

        def check(window: DocxWindow) -> bool:
            if not window.document.has_document:
                return GLib.SOURCE_CONTINUE
            adjustment = window.document.widget.get_vadjustment()
            maximum = adjustment.get_upper() - adjustment.get_lower() - adjustment.get_page_size()
            if maximum <= 40:
                # The document object exists before GTK has finished measuring
                # its replacement page stack; wait for a real scroll range.
                return GLib.SOURCE_CONTINUE

            if not state["replaced"]:
                if window._path_status.get_text() != _compact_path(watched.resolve()):
                    return finish(window, "the bottom-left status must show the document realpath")
                if not window._page_indicator.get_text().endswith(
                    f" / {window.document.page_count}"
                ):
                    return finish(window, "the bottom-right status must show the page count")
                # A multi-page document starts at its first readable position,
                # rather than appearing fully read before its first scroll.
                if window._reading_progress.get_fraction() > 0.01:
                    return finish(
                        window,
                        "the reading-progress rule must begin at zero for a scrollable document",
                    )
                progress_rule = window._reading_progress.get_parent()
                status_area = progress_rule.get_parent()
                if progress_rule.get_allocated_height() > 3:
                    return finish(
                        window,
                        "search markers must not expand the three-pixel status rule",
                    )
                if status_area.get_allocated_height() > 35:
                    return finish(
                        window,
                        "search markers must not expand the compact status area",
                    )
                document_layout = window.document.widget.get_parent().get_parent()
                workspace = document_layout.get_parent()
                root = workspace.get_parent()
                if (
                    root.get_allocated_height()
                    - status_area.get_allocated_height()
                    - workspace.get_allocated_height()
                    > 3
                ):
                    return finish(
                        window,
                        "the document canvas must fill the workspace above the status bar",
                    )
                _horizontal_policy, vertical_policy = window.document.widget.get_policy()
                if vertical_policy != Gtk.PolicyType.AUTOMATIC:
                    return finish(
                        window,
                        "the document viewport must retain its vertical scrolling model",
                    )
                if window.get_titlebar() is not None:
                    return finish(window, "the document window must not render a top bar")
                # A print preview must use separately allocated pages, not a
                # seamless document-height canvas.  Check GTK's real layout,
                # rather than a unit-test approximation of it.
                pages = window.document._page_geometries()
                if len(pages) < 2:
                    return finish(window, f"expected a multi-page preview; pages={pages}")
                gaps = [
                    following.top - preceding.bottom
                    for preceding, following in zip(pages, pages[1:])
                ]
                if any(gap < PAGE_GAP for gap in gaps):
                    return finish(
                        window,
                        f"pages are not visibly separated; expected {PAGE_GAP}px gaps, got {gaps}",
                    )
                window.document.go_to_page(0)
                if not window._on_key_press(
                    window, SimpleNamespace(keyval=Gdk.KEY_J, state=0)
                ) or window.document.current_page_index != 1:
                    return finish(window, "J must advance one document page")
                expected_progress = (adjustment.get_value() - adjustment.get_lower()) / maximum
                if window._reading_progress_source == 0:
                    return finish(window, "the reading-progress rule must animate toward document scrolling")
                if not 0.0 <= window._reading_progress.get_fraction() <= expected_progress:
                    return finish(window, "the reading-progress rule must follow document scrolling")
                if not window._on_key_press(
                    window, SimpleNamespace(keyval=Gdk.KEY_K, state=0)
                ) or window.document.current_page_index != 0:
                    return finish(window, "K must return to the previous document page")
                if not window.document.outline:
                    return finish(window, "expected the fixture headings in the document outline")
                if not window._on_key_press(
                    window, SimpleNamespace(keyval=Gdk.KEY_Tab, state=0)
                ) or not window._outline_panel.get_visible():
                    return finish(window, "Tab must open the document outline")
                if window._outline_panel.get_parent() is not window._document_layout:
                    return finish(window, "opening the outline must reserve document layout space")
                title, _scroller, _empty = window._outline_panel.get_children()
                if not title.get_visible() or not window._outline_tree.get_visible():
                    return finish(window, "Tab must show the outline title and heading list")
                reference_column = window._outline_tree.get_column(0)
                title_column = window._outline_tree.get_column(1)
                if reference_column is None or title_column is None:
                    return finish(window, "outline must show reference and heading columns")
                (reference_renderer,) = reference_column.get_cells()
                (renderer,) = title_column.get_cells()
                if (
                    window._outline_tree.get_expander_column() is not title_column
                    or reference_renderer.get_property("scale") >= 0.9
                    or reference_renderer.get_property("xalign") != 1.0
                    or window._outline_tree.get_margin_start() != OUTLINE_NAV_SPACING
                    or reference_column.get_sizing()
                    != Gtk.TreeViewColumnSizing.FIXED
                    # Keep the visual gutter to three units: one beside a
                    # single-digit reference and one before the expander.
                    or reference_column.get_fixed_width() != OUTLINE_NAV_SPACING * 3
                ):
                    return finish(
                        window,
                        "outline references must occupy a fixed left gutter before tree arrows",
                    )
                if (
                    renderer.get_property("ellipsize") != Pango.EllipsizeMode.END
                    or renderer.get_property("width-chars") > 22
                ):
                    return finish(window, "outline headings must truncate early with an ellipsis")
                if window.document.outline and window._outline_empty.get_visible():
                    return finish(window, "an outline with headings must not show the empty-state message")
                selected = window._selected_outline_index()
                if selected is None:
                    return finish(window, "opening the document outline must select a heading")
                selected_path = window._outline_row_paths[selected]
                if not window._on_key_press(
                    window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_l, state=0)
                ) or not window._outline_tree.row_expanded(selected_path):
                    return finish(window, "l must expand the selected nested heading")
                if not window._on_key_press(
                    window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_h, state=0)
                ) or window._outline_tree.row_expanded(selected_path):
                    return finish(window, "h must collapse the selected nested heading")
                if not window._on_key_press(
                    window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_l, state=0)
                ) or not window._outline_tree.row_expanded(selected_path):
                    return finish(window, "l must expand the selected nested heading")
                expected_index = min(selected + 1, len(window.document.outline) - 1)
                selected_entry = window.document.outline[expected_index]
                heading_page = selected_entry.page_index
                animation_frames: list[float] = []
                original_show_locator = window.document._show_outline_locator
                shown_pages: list[object] = []
                original_draw_locators: list[tuple[object, object]] = []

                def record_show_locator(shown_page: object, top: float) -> None:
                    shown_pages.append(shown_page)
                    original_draw_locator = shown_page._draw_outline_locator

                    def record_outline_locator(context: object) -> None:
                        animation_frames.append(shown_page._outline_locator_elapsed_ms)
                        original_draw_locator(context)

                    original_draw_locators.append((shown_page, original_draw_locator))
                    shown_page._draw_outline_locator = record_outline_locator
                    original_show_locator(shown_page, top)

                window.document._show_outline_locator = record_show_locator
                if not window._on_key_press(
                    window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_j, state=0)
                ) or window._selected_outline_index() != expected_index:
                    return finish(window, "j must select the next outline heading")
                if not window._on_key_press(
                    window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_Return, state=0)
                ):
                    return finish(window, "Enter must jump to the selected outline heading")
                frame_loop = GLib.MainLoop()
                GLib.timeout_add(
                    OUTLINE_LOCATOR_ARRIVAL_MS + 120,
                    lambda: (frame_loop.quit(), GLib.SOURCE_REMOVE)[1],
                )
                frame_loop.run()
                window.document._show_outline_locator = original_show_locator
                for observed_page, original_draw_locator in original_draw_locators:
                    observed_page._draw_outline_locator = original_draw_locator
                page = shown_pages[0] if shown_pages else None
                if len(animation_frames) < 3 or not any(
                    0 < elapsed < OUTLINE_LOCATOR_ARRIVAL_MS
                    for elapsed in animation_frames
                ):
                    return finish(
                        window,
                        "outline arrival did not repaint intermediate frames: "
                        f"{animation_frames}; current={getattr(page, '_outline_locator_elapsed_ms', None)}; "
                        f"source={window.document._outline_locator_source}; "
                        f"target-page={window.document._pages.index(page) if page in window.document._pages else None}",
                    )
                if window.document.current_page_index != heading_page:
                    return finish(window, "outline jump must move to the selected heading page")
                if arguments.outline_count_only:
                    # Use a collapsed child in a live TreeView so this follows
                    # exactly the menu's visible-row navigation semantics.
                    window._toggle_outline()
                    window.document._outline = [
                        OutlineEntry("Chapter one", 0, None, 0),
                        OutlineEntry("Hidden section", 0, None, 1),
                        *(
                            OutlineEntry(f"Chapter {number}", 0, None, 0)
                            for number in range(2, 15)
                        ),
                    ]
                    window._update_outline()
                    window._toggle_outline()
                    window._select_outline_entry(0)

                    def reference_for(index: int) -> str:
                        row = window._outline_store.get_iter(window._outline_row_paths[index])
                        return window._outline_store.get_value(row, 0)

                    if not window._on_key_press(
                        window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_l, state=0)
                    ) or {
                        index: reference_for(index)
                        for index in (0, 1, 2)
                    } != {0: "0", 1: "1", 2: "2"}:
                        return finish(
                            window,
                            "expanding an outline branch must number its visible headings",
                        )
                    if not window._on_key_press(
                        window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_h, state=0)
                    ):
                        return finish(window, "h must collapse the synthetic outline branch")
                    if {
                        index: reference_for(index)
                        for index in (0, 1, 2, 5)
                    } != {0: "0", 1: "", 2: "1", 5: "4"}:
                        return finish(
                            window,
                            "outline references must number visible headings relative to selection",
                        )
                    if not window._on_key_press(
                        window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_4, state=0)
                    ) or not window._on_key_press(
                        window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_j, state=0)
                    ) or window._selected_outline_index() != 5:
                        return finish(
                            window,
                            "4j must move four visible outline rows",
                        )
                    if {
                        index: reference_for(index)
                        for index in (0, 5, 6)
                    } != {0: "-4", 5: "0", 6: "1"}:
                        return finish(
                            window,
                            "outline references must follow the selected heading",
                        )
                    if not window._on_key_press(
                        window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_4, state=0)
                    ) or not window._on_key_press(
                        window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_k, state=0)
                    ) or window._selected_outline_index() != 0:
                        return finish(
                            window,
                            "4k must move four visible outline rows backward",
                        )
                    if not window._on_key_press(
                        window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_1, state=0)
                    ) or not window._on_key_press(
                        window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_2, state=0)
                    ) or not window._on_key_press(
                        window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_j, state=0)
                    ) or window._selected_outline_index() != 13:
                        return finish(
                            window,
                            "12j must move twelve visible outline rows",
                        )
                    print("count-prefixed outline navigation works")
                    return finish(window)
                if arguments.scroll_navigation_only:
                    print("scroll and outline navigation work")
                    return finish(window)
                if selected_entry.top is not None and not arguments.search_cancel_only:
                    page = window.document._pages[heading_page]
                    destination = page.destination_y(selected_entry.top)
                    expected_target = min(
                        max(destination - adjustment.get_page_size() * 0.30, 0.0),
                        window.document._maximum_scroll(),
                    )
                    actual_target = adjustment.get_value() - adjustment.get_lower()
                    if abs(actual_target - expected_target) > 1:
                        return finish(window, "outline jump must keep the heading in viewport context")
                    if page._outline_locator_top != selected_entry.top:
                        return finish(window, "outline jump must briefly mark the heading location")
                if arguments.outline_animation_only:
                    print(f"outline animation frames: {animation_frames}")
                    return finish(window)
                if not window._on_key_press(
                    window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_Tab, state=0)
                ) or window._outline_panel.get_visible():
                    return finish(window, "Tab must close the document outline")
                # When the adaptive limit collapses child headings, j/k must
                # move among the rows a reader can actually see.
                window.document._outline = [
                    OutlineEntry("Chapter one", 0, None, 0),
                    OutlineEntry("Hidden section", 0, None, 1),
                    *(OutlineEntry(f"Chapter {number}", 0, None, 0) for number in range(2, 10)),
                ]
                window._update_outline()
                window._toggle_outline()
                if not window._on_key_press(
                    window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_j, state=0)
                ) or window._selected_outline_index() != 2:
                    return finish(window, "j must skip collapsed outline children")
                # A Vim-style count must apply to the navigable tree rows, so
                # 4j skips four visible headings rather than four model rows.
                window._select_outline_entry(0)
                if not window._on_key_press(
                    window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_4, state=0)
                ) or not window._on_key_press(
                    window._outline_tree, SimpleNamespace(keyval=Gdk.KEY_j, state=0)
                ) or window._selected_outline_index() != 5:
                    return finish(window, "4j must select the fourth visible outline heading")
                window._toggle_outline()
                if not window._on_key_press(
                    window, SimpleNamespace(keyval=Gdk.KEY_slash, state=0)
                ) or not window._search_panel.get_visible():
                    return finish(window, "/ must open the document search prompt")
                if (
                    window._search_panel.get_valign() != Gtk.Align.CENTER
                    or window._search_panel.get_margin_bottom() != 0
                    or isinstance(window._search_entry, Gtk.SearchEntry)
                    or window._search_panel.get_parent().get_child()
                    is not window.document.widget
                ):
                    return finish(
                        window,
                        "search must be a centered plain-entry command prompt",
                    )
                # Search input is insert mode: ordinary printable keys must
                # reach its entry instead of invoking viewer navigation.
                if window._on_key_press(
                    window._search_entry, SimpleNamespace(keyval=Gdk.KEY_j, state=0)
                ):
                    return finish(window, "search input must not claim ordinary typing as j navigation")
                window._search_entry.set_text("purpose")
                if len(window._search_matches) < 2:
                    return finish(window, "search must find repeated PDF text")
                expected_search_status = (
                    f"{window._search_index + 1} of {len(window._search_matches)}"
                )
                marker_fractions = getattr(
                    window, "_search_match_marker_fractions", lambda: []
                )()
                if (
                    window._search_status.get_text() != expected_search_status
                    or len(marker_fractions) != len(window._search_matches)
                    or any(not 0.0 <= fraction <= 1.0 for fraction in marker_fractions)
                ):
                    return finish(
                        window,
                        "search must show current match count and document-wide match markers",
                    )
                active_match = window._search_matches[window._search_index]
                if window.document._pages[active_match.page_index]._search_highlight != active_match:
                    return finish(window, "search must visibly highlight the active PDF text match")
                search_target = min(
                    max(
                        window.document._pages[active_match.page_index].destination_y(
                            active_match.top
                        )
                        - adjustment.get_page_size() * 0.50,
                        0.0,
                    ),
                    window.document._maximum_scroll(),
                )
                if abs(adjustment.get_value() - adjustment.get_lower() - search_target) > 1:
                    return finish(window, "search must center the active match when space permits")
                initial_match = window._search_index
                if not window._on_search_key_press(
                    window._search_entry, SimpleNamespace(keyval=Gdk.KEY_Return, state=0)
                ) or window._search_index == initial_match or window._search_panel.get_visible():
                    return finish(window, "Enter must commit and close the document search prompt")
                committed_match = window._search_index
                # Enter reveals a previously hidden status label. Let GTK
                # complete the following layout frame before checking its
                # centred footer allocation.
                search_status_layout = GLib.MainLoop()
                GLib.timeout_add(
                    20,
                    lambda: (search_status_layout.quit(), GLib.SOURCE_REMOVE)[1],
                )
                search_status_layout.run()
                search_session_status = getattr(
                    window, "_search_session_status", None
                )
                marker_fractions = window._search_match_marker_fractions()
                expected_session_status = (
                    f"Search · {committed_match + 1} of "
                    f"{len(window._search_matches)}"
                )
                search_state_problems: list[str] = []
                if search_session_status is None:
                    search_state_problems.append("missing persistent search status")
                else:
                    status_bar = search_session_status.get_parent()
                    search_allocation = search_session_status.get_allocation()
                    if (
                        not search_session_status.get_visible()
                        or search_session_status.get_text() != expected_session_status
                        or status_bar is None
                        or abs(
                            search_allocation.x
                            + search_allocation.width / 2
                            - status_bar.get_allocation().width / 2
                        )
                        > 1
                    ):
                        search_state_problems.append(
                            "search status is not centred in the footer: "
                            f"visible={search_session_status.get_visible()}; "
                            f"text={search_session_status.get_text()!r}; "
                            f"search=({search_allocation.x},{search_allocation.y},"
                            f"{search_allocation.width},{search_allocation.height}); "
                            f"footer_width={status_bar.get_allocation().width if status_bar else None}"
                        )
                if (
                    not marker_fractions
                    or abs(
                        window._reading_progress.get_fraction()
                        - marker_fractions[committed_match]
                    )
                    > 0.01
                ):
                    search_state_problems.append(
                        "progress fill and active search marker disagree: "
                        f"progress={window._reading_progress.get_fraction():.4f}; "
                        f"marker={marker_fractions[committed_match] if marker_fractions else None}; "
                        f"animation={window._reading_progress_source}"
                    )
                if search_state_problems:
                    return finish(window, "; ".join(search_state_problems))
                if not window._on_key_press(
                    window, SimpleNamespace(keyval=Gdk.KEY_n, state=0)
                ) or window._search_index == committed_match:
                    return finish(window, "n must advance after a committed search")
                if window._search_status.get_text() != (
                    f"{window._search_index + 1} of {len(window._search_matches)}"
                ):
                    return finish(window, "n must update the visible search position")
                if search_session_status.get_text() != (
                    f"Search · {window._search_index + 1} of "
                    f"{len(window._search_matches)}"
                ):
                    return finish(window, "n must update the persistent search position")
                marker_fractions = window._search_match_marker_fractions()
                if abs(
                    window._reading_progress.get_fraction()
                    - marker_fractions[window._search_index]
                ) > 0.01:
                    return finish(window, "n must keep progress on the active search marker")
                if not window._on_key_press(
                    window, SimpleNamespace(keyval=Gdk.KEY_N, state=0)
                ) or window._search_index != committed_match:
                    return finish(window, "N must return to the previous search match")
                if window._search_status.get_text() != (
                    f"{window._search_index + 1} of {len(window._search_matches)}"
                ):
                    return finish(window, "N must update the visible search position")
                if search_session_status.get_text() != (
                    f"Search · {window._search_index + 1} of "
                    f"{len(window._search_matches)}"
                ):
                    return finish(window, "N must update the persistent search position")
                marker_fractions = window._search_match_marker_fractions()
                if abs(
                    window._reading_progress.get_fraction()
                    - marker_fractions[window._search_index]
                ) > 0.01:
                    return finish(window, "N must keep progress on the active search marker")
                # Escape after committing a search must remove every part of
                # the search session, not merely hide the prompt.
                if not window._on_key_press(
                    window, SimpleNamespace(keyval=Gdk.KEY_Escape, state=0)
                ):
                    return finish(window, "Escape must cancel the active search")
                if (
                    window._search_panel.get_visible()
                    or window._search_matches
                    or window._search_index != -1
                    or search_session_status.get_visible()
                    or any(page._search_highlight is not None for page in window.document._pages)
                ):
                    return finish(window, "Escape must clear the search highlight and results")
                if window._on_key_press(
                    window, SimpleNamespace(keyval=Gdk.KEY_n, state=0)
                ) or window._on_key_press(
                    window, SimpleNamespace(keyval=Gdk.KEY_N, state=0)
                ):
                    return finish(window, "n/N must be inactive after Escape cancels search")
                if arguments.search_cancel_only:
                    print("Escape clears the committed search session")
                    return finish(window)
                if not window._on_key_press(
                    window, SimpleNamespace(keyval=Gdk.KEY_colon, state=0)
                ) or not window._page_jump_panel.get_visible():
                    return finish(window, ": must open the page jump prompt")
                window._page_jump_entry.set_text("3")
                if not window._on_page_jump_key_press(
                    window._page_jump_entry,
                    SimpleNamespace(keyval=Gdk.KEY_Return, state=0),
                ) or window.document.current_page_index != 2:
                    return finish(window, ":3 must jump to page three")
                if window._page_jump_panel.get_visible():
                    return finish(window, "committing a page jump must close its prompt")
                adjustment.set_value(adjustment.get_lower() + maximum * 0.62)
                state["expected"] = window.document.capture_position()
                state["initial_revision"] = window._rendered_revision
                state["before"] = (
                    adjustment.get_value(),
                    maximum,
                    window.document._page_geometries(),
                )
                replacement = watched.with_name("replacement.docx")
                shutil.copy2(watched, replacement)
                replacement.replace(watched)
                state["replaced"] = True
                return GLib.SOURCE_CONTINUE

            if window._rendered_revision == state["initial_revision"]:
                return GLib.SOURCE_CONTINUE
            if (
                window._process is not None
                or window.document.restore_pending
            ):
                return GLib.SOURCE_CONTINUE

            expected = state["expected"]
            actual = window.document.capture_position()
            if expected is None or expected.page_index != actual.page_index:
                return finish(
                    window,
                    f"page changed from {expected} to {actual}; before={state['before']}",
                )
            if abs((expected.page_fraction or 0.0) - (actual.page_fraction or 0.0)) > 0.03:
                return finish(
                    window,
                    "within-page offset changed from " f"{expected} to {actual}; before={state['before']}",
                )
            print(f"position preserved: {actual}")
            return finish(window)

        def activate(_application: Gtk.Application) -> None:
            window = DocxWindow(application, watched)
            window.show_all()
            GLib.timeout_add(100, check, window)
            GLib.timeout_add(
                TIMEOUT_MS,
                lambda: finish(
                    window,
                    "timed out waiting for the live refresh "
                    f"(revision={window._revision}; status={window._last_status!r})",
                ),
            )

        application.connect("activate", activate)
        application.run(["see-docx-ui-smoke"])
        if error := result.get("error"):
            print(f"FAIL: {error}")
            return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
