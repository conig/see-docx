"""GTK/Poppler document window with live DOCX refresh."""

from __future__ import annotations

import ctypes
import ctypes.util
from difflib import SequenceMatcher
from functools import lru_cache
from html import escape as html_escape
import json
import math
import os
import shlex
from collections.abc import Callable, Iterable
from dataclasses import dataclass
from datetime import datetime
from pathlib import Path
import zipfile
import xml.etree.ElementTree as ET

import cairo
import gi

gi.require_version("Gdk", "3.0")
gi.require_version("Gtk", "3.0")
gi.require_version("Poppler", "0.18")

from gi.repository import Gdk, Gio, GLib, Gtk, Pango, Poppler

from .converter import (
    ConversionError,
    ConversionPaths,
    LibreOfficeConverter,
    PandocConversionPaths,
    PandocConverter,
)
from .position import (
    DocumentPosition,
    PageGeometry,
    capture_position,
    page_index_at_scroll,
    restore_position,
)

APPLICATION_ID = "io.github.conig.seedocx"
DEFAULT_ZOOM = 1.25
MIN_ZOOM = 0.60
MIN_FIT_ZOOM = 0.10
MAX_ZOOM = 2.00
ZOOM_STEP = 0.10
REFRESH_DEBOUNCE_MS = 450
SCROLL_STEP = 56
TEXT_SELECTION_LINE_PADDING = 4.0
SELECTION_AUTO_SCROLL_TICK_MS = 16
SELECTION_AUTO_SCROLL_MIN_STEP = 12.0
SELECTION_AUTO_SCROLL_MAX_STEP = 52.0
_THEME_STATE_SCHEMA_VERSION = 1
_THEME_STATE_ROLES = frozenset(
    {
        "background_alt",
        "canvas",
        "surface",
        "surface_raised",
        "foreground",
        "foreground_muted",
        "metadata",
        "gtk_command",
        "highlight",
        "dim",
        "selection",
        "selection_foreground",
        "separator",
    }
)

PAGE_GAP = 28
PAGE_MARGIN = 36
COMMENTS_PANEL_MIN_WIDTH = 300
COMMENTS_PANEL_MAX_WIDTH = 420
COMMENTS_PANEL_FRACTION = 0.34
COMMENTS_DOCUMENT_MIN_WIDTH = 520
COMMENTS_PANEL_GUTTER = 16
COMMENTS_HEADER_HEIGHT = 50
COMMENT_CARD_GAP = 12
COMMENT_VIEWPORT_BUFFER = 120
COMMENT_INACTIVE_BODY_HEIGHT = 76
COMMENT_BODY_MIN_HEIGHT = 28
COMMENT_ACTIVE_BODY_MAX_HEIGHT = 320
# Keep the right edge of every card fixed. The inactive left inset leaves the
# active card roughly 20% wider without moving either card's right edge.
COMMENT_INACTIVE_CARD_INSET = 48
OUTLINE_HEADING_MAX_CHARS = 22
OUTLINE_NAV_SPACING = 7
OUTLINE_REFERENCE_MARGIN_WIDTH = OUTLINE_NAV_SPACING * 3
OUTLINE_TITLE_PADDING_START = OUTLINE_NAV_SPACING * 2
OUTLINE_CONTENT_MARGIN_START = OUTLINE_NAV_SPACING
OUTLINE_INITIAL_VISIBLE_LIMIT = 9
OUTLINE_CONTEXT_VIEWPORT_FRACTION = 0.30
SEARCH_CONTEXT_VIEWPORT_FRACTION = 0.50
OUTLINE_LOCATOR_DURATION_MS = 1_583
OUTLINE_LOCATOR_TICK_MS = 16
OUTLINE_LOCATOR_ARRIVAL_MS = 583
OUTLINE_LOCATOR_BLOOM_MS = 750
OUTLINE_LOCATOR_RIPPLE_EXPANSION_MS = 750
OUTLINE_LOCATOR_RIPPLE_FADE_START_MS = 400
OUTLINE_LOCATOR_FALLBACK_HEIGHT = 34.0
READING_PROGRESS_DURATION_MS = 180
READING_PROGRESS_TICK_MS = 16
HINT_CHARS = "asdfghjkl"
_WORDPROCESSINGML = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
_WORD_2010_WORDML = "{http://schemas.microsoft.com/office/word/2010/wordml}"
_WORD_2012_WORDML = "{http://schemas.microsoft.com/office/word/2012/wordml}"
EXPORT_FORMATS = ("PDF", "Plain text")
EXPORT_FORMAT_DESCRIPTIONS = {
    "PDF": "Portable Document Format",
    "Plain text": "UTF-8 text via Pandoc",
}


def _style(widget: Gtk.Widget, class_name: str) -> None:
    widget.get_style_context().add_class(class_name)


def _label(text: str, *, xalign: float = 0.0) -> Gtk.Label:
    return Gtk.Label(label=text, xalign=xalign)


class _FooterStatusBar(Gtk.Overlay):
    """A footer whose transient status can remain centred independently."""

    def __init__(self) -> None:
        super().__init__()
        self._centred_overlay: Gtk.Widget | None = None
        self._centred_layer: Gtk.Fixed | None = None

    def add_centred_overlay(self, widget: Gtk.Widget) -> None:
        """Add *widget* above the footer content, centred within this footer."""

        self._centred_overlay = widget
        self._centred_layer = Gtk.Fixed()
        self._centred_layer.set_halign(Gtk.Align.FILL)
        self._centred_layer.set_valign(Gtk.Align.FILL)
        self._centred_layer.set_hexpand(True)
        self._centred_layer.set_vexpand(True)
        self._centred_layer.connect(
            "size-allocate", self._on_centred_layer_size_allocate
        )
        widget.connect("show", self._on_centred_overlay_show)
        self.add_overlay(self._centred_layer)
        self.set_overlay_pass_through(self._centred_layer, True)
        self._centred_layer.put(widget, 0, 0)
        self._centre_overlay()

    def _on_centred_layer_size_allocate(
        self, _layer: Gtk.Fixed, _allocation: Gdk.Rectangle
    ) -> None:
        self._centre_overlay()

    def _on_centred_overlay_show(self, _widget: Gtk.Widget) -> None:
        self._centre_overlay()

    def _centre_overlay(self) -> None:
        if self._centred_layer is None or self._centred_overlay is None:
            return
        width = self._centred_overlay.get_preferred_width()[1]
        height = self._centred_overlay.get_preferred_height()[1]
        self._centred_layer.move(
            self._centred_overlay,
            max((self._centred_layer.get_allocated_width() - width) // 2, 0),
            max((self._centred_layer.get_allocated_height() - height) // 2, 0),
        )


def _compact_path(path: Path, home: Path | None = None) -> str:
    """Display paths under the user's home directory with a familiar ``~``."""

    home = Path.home() if home is None else home
    try:
        relative = path.relative_to(home)
    except ValueError:
        return str(path)
    return "~" if relative == Path(".") else f"~/{relative}"


def _pdf_export_path(path: Path) -> Path:
    """Ensure the destination selected for a PDF export has a PDF suffix."""

    if path.suffix.lower() == ".pdf":
        return path
    return path.with_name(f"{path.name}.pdf")


def _plain_text_export_path(path: Path) -> Path:
    """Ensure the destination selected for plain text has a TXT suffix."""

    if path.suffix.lower() == ".txt":
        return path
    return path.with_name(f"{path.name}.txt")


def _number_to_hint(number: int, width: int) -> str:
    """Return qutebrowser's base-N hint string for *number*."""

    base = len(HINT_CHARS)
    characters: list[str] = []
    while True:
        remainder = number % base
        characters.insert(0, HINT_CHARS[remainder])
        number = (number - remainder) // base
        if number <= 0:
            break
    return "".join(characters).rjust(width, HINT_CHARS[0])


def hint_codes(count: int) -> list[str]:
    """Generate See Mail-compatible, prefix-free home-row hint codes."""

    if count <= 0:
        return []

    base = len(HINT_CHARS)
    width = 1
    while base**width < count:
        width += 1

    short_count = 0
    if width > 1:
        short_count = (base**width - count) // (base - 1)
    long_count = count - short_count

    codes = [_number_to_hint(number, width - 1) for number in range(short_count)]
    start = short_count * base
    codes.extend(
        _number_to_hint(number, width)
        for number in range(start, start + long_count)
    )

    buckets = [[] for _character in HINT_CHARS]
    for index, code in enumerate(codes):
        buckets[index % base].append(code)
    return [code for bucket in buckets for code in bucket]


def _fit_zoom_for_viewport(
    page_width: float, page_height: float, viewport_width: float, viewport_height: float
) -> float:
    """Return the zoom that keeps a complete page and its margins in view."""

    if page_width <= 0 or page_height <= 0:
        return MIN_FIT_ZOOM
    width_zoom = (viewport_width - 2 * PAGE_MARGIN) / page_width
    height_zoom = (viewport_height - 2 * PAGE_MARGIN) / page_height
    return max(MIN_FIT_ZOOM, min(width_zoom, height_zoom))


def _lookup_color(widget: Gtk.Widget, name: str, fallback: str) -> str:
    found, color = widget.get_style_context().lookup_color(name)
    return color.to_string() if found else fallback


def _theme_state_path() -> Path:
    """Return the documented SC1 Command UI state API location."""

    state_home = os.environ.get("XDG_STATE_HOME")
    if state_home:
        return Path(state_home) / "sc1-command-ui" / "current.json"
    return Path.home() / ".local" / "state" / "sc1-command-ui" / "current.json"


def _is_hex_color(value: object) -> bool:
    return (
        isinstance(value, str)
        and len(value) == 7
        and value.startswith("#")
        and all(character in "0123456789abcdefABCDEF" for character in value[1:])
    )


@lru_cache(maxsize=1)
def _current_theme_state_roles() -> dict[str, str] | None:
    """Read the complete SC1 palette snapshot, or decline it as unusable.

    The state file is atomically published by SC1 Command UI. A complete
    snapshot keeps this app from mixing a partially upgraded API with GTK
    fallback values. GTK remains the compatibility path before first apply.
    """

    try:
        document = json.loads(_theme_state_path().read_text(encoding="utf-8"))
    except (OSError, UnicodeDecodeError, json.JSONDecodeError):
        return None

    if (
        not isinstance(document, dict)
        or document.get("schema_version") != _THEME_STATE_SCHEMA_VERSION
        or not isinstance(document.get("variant"), str)
        or not isinstance(document.get("roles"), dict)
    ):
        return None

    roles = document["roles"]
    if not _THEME_STATE_ROLES.issubset(roles) or not all(
        _is_hex_color(roles[role]) for role in _THEME_STATE_ROLES
    ):
        return None
    return {role: roles[role] for role in _THEME_STATE_ROLES}


def _theme_palette_from_state(roles: dict[str, str]) -> dict[str, str]:
    """Map stable theme roles to See DOCX's visual vocabulary."""

    return {
        "background": roles["background_alt"],
        "canvas": roles["canvas"],
        "panel": roles["surface_raised"],
        "panel_dark": roles["surface"],
        "foreground": roles["foreground"],
        "view_foreground": roles["gtk_command"],
        "text": roles["metadata"],
        "muted": roles["foreground_muted"],
        "metadata": roles["metadata"],
        "accent": roles["gtk_command"],
        "highlight": roles["highlight"],
        "accent_dim": roles["dim"],
        "selected_background": roles["selection"],
        "selected_foreground": roles["selection_foreground"],
        "separator": roles["separator"],
    }


def _theme_palette(widget: Gtk.Widget) -> dict[str, str]:
    """Prefer the SC1 state API, then resolve semantic GTK fallbacks."""

    if roles := _current_theme_state_roles():
        return _theme_palette_from_state(roles)

    background = _lookup_color(widget, "theme_bg_color", "#202326")
    canvas = _lookup_color(widget, "theme_base_color", background)
    foreground = _lookup_color(widget, "theme_fg_color", "#e7ebf3")
    selected_background = _lookup_color(widget, "theme_selected_bg_color", "#26364a")
    accent = _lookup_color(
        widget,
        "sc1-command-green",
        _lookup_color(widget, "success_color", "#65d48a"),
    )
    return {
        "background": _lookup_color(widget, "sc1-bg-1", background),
        # The document sheet needs a clearly contrasting surround. GTK's
        # generic base colour is often white, which otherwise erases its edge.
        "canvas": _lookup_color(widget, "sc1-canvas-bg", "#30363d"),
        "panel": _lookup_color(widget, "sc1-bg-panel-raised", canvas),
        "panel_dark": _lookup_color(widget, "sc1-bg-panel", background),
        "foreground": _lookup_color(widget, "sc1-fg-normal", foreground),
        "view_foreground": _lookup_color(widget, "sc1-command-green", accent),
        "text": _lookup_color(widget, "theme_text_color", foreground),
        "muted": _lookup_color(widget, "sc1-fg-muted", "#a9b2bd"),
        "metadata": _lookup_color(widget, "sc1-terran-blue", "#9DB6E9"),
        "accent": accent,
        "highlight": _lookup_color(widget, "sc1-command-highlight", foreground),
        "accent_dim": _lookup_color(widget, "sc1-command-dim", accent),
        "selected_background": _lookup_color(
            widget, "sc1-selection-bg-solid", selected_background
        ),
        "selected_foreground": _lookup_color(
            widget,
            "sc1-selection-fg",
            _lookup_color(widget, "sc1-command-highlight", foreground),
        ),
        "separator": _lookup_color(
            widget, "sc1-separator", _lookup_color(widget, "borders", "#3b4148")
        ),
    }


def _rgb(color_spec: str) -> tuple[float, float, float]:
    """Resolve a GTK colour string to the RGB components Cairo expects."""

    color = Gdk.RGBA()
    if not color.parse(color_spec):
        return (0.0, 0.0, 0.0)
    return color.red, color.green, color.blue


def _coalesce_comment_rectangles(
    rectangles: Iterable[tuple[float, float, float, float]],
) -> tuple[tuple[float, float, float, float], ...]:
    """Merge repeated or adjacent glyph marks on the same rendered line."""

    lines: dict[float, list[tuple[float, float, float, float]]] = {}
    for rectangle in rectangles:
        lines.setdefault(round(rectangle[1], 1), []).append(rectangle)

    merged: list[tuple[float, float, float, float]] = []
    for line in lines.values():
        current: list[float] | None = None
        for left, top, right, bottom in sorted(line, key=lambda item: item[0]):
            if current is None or left > current[2] + 0.5:
                if current is not None:
                    merged.append(tuple(current))
                current = [left, top, right, bottom]
                continue
            current[0] = min(current[0], left)
            current[1] = min(current[1], top)
            current[2] = max(current[2], right)
            current[3] = max(current[3], bottom)
        if current is not None:
            merged.append(tuple(current))
    return tuple(merged)


def _app_css(widget: Gtk.Widget) -> bytes:
    palette = _theme_palette(widget)
    return f"""
window.see-docx-window {{
  background-color: {palette["background"]};
  color: {palette["foreground"]};
}}
.see-docx-root {{ background-color: {palette["background"]}; }}
.see-docx-page-indicator {{
  color: {palette["metadata"]};
  font-size: 0.87em;
}}
.see-docx-workspace,
.see-docx-workspace viewport {{ background-color: {palette["canvas"]}; }}
.see-docx-workspace scrollbar.vertical,
.see-docx-workspace scrollbar.vertical slider {{
  min-width: 0;
  min-height: 0;
  margin: 0;
  padding: 0;
  border: 0;
  opacity: 0;
}}
.see-docx-pages {{ background-color: {palette["canvas"]}; }}
.see-docx-page {{
  background-color: #ffffff;
  border: 1px solid {palette["separator"]};
  box-shadow: 0 5px 20px rgba(0, 0, 0, 0.32);
}}
.see-docx-outline {{
  background-color: {palette["panel_dark"]};
  border-right: 1px solid {palette["separator"]};
}}
.see-docx-outline-title {{
  color: {palette["muted"]};
  font-size: 0.80em;
  font-weight: 700;
  letter-spacing: 0.08em;
  padding: 12px {OUTLINE_TITLE_PADDING_START}px 8px;
}}
.see-docx-outline treeview.view {{
  background-color: {palette["panel_dark"]};
  color: {palette["view_foreground"]};
  font-size: 0.96em;
  -GtkTreeView-horizontal-separator: {OUTLINE_NAV_SPACING}px;
  -GtkTreeView-level-indentation: {OUTLINE_NAV_SPACING}px;
}}
.see-docx-outline treeview.view.expander {{
  color: {palette["muted"]};
}}
.see-docx-outline treeview.view.expander:checked,
.see-docx-outline treeview.view.expander:hover,
.see-docx-outline treeview.view.expander:active {{
  color: {palette["highlight"]};
}}
.see-docx-outline treeview.view:selected,
.see-docx-outline treeview.view:selected:focus {{
  box-shadow: none;
  color: {palette["selected_foreground"]};
}}
.see-docx-outline-empty {{
  color: {palette["muted"]};
  padding: 8px 14px;
}}
.see-docx-comments {{
  background-color: {palette["panel_dark"]};
  border-left: 1px solid {palette["separator"]};
}}
.see-docx-comments-header {{
  min-height: {COMMENTS_HEADER_HEIGHT}px;
  padding: 7px 12px 6px;
}}
.see-docx-comments-header-top {{
  min-height: 17px;
}}
.see-docx-comments-header-meta {{
  min-height: 14px;
}}
.see-docx-comments scrolledwindow,
.see-docx-comments scrolledwindow viewport {{
  background-color: {palette["panel_dark"]};
  border: 0;
}}
.see-docx-comments scrolledwindow scrollbar.vertical slider {{
  min-width: 5px;
  border-radius: 3px;
  background-color: {palette["separator"]};
}}
.see-docx-comments-title {{
  color: {palette["muted"]};
  font-size: 0.80em;
  font-weight: 700;
  letter-spacing: 0.10em;
}}
.see-docx-comments.comments-focused .see-docx-comments-title {{
  color: {palette["highlight"]};
}}
.see-docx-comments.comments-focused .see-docx-comments-header {{
  border-bottom: 1px solid {palette["accent"]};
}}
.see-docx-comments.comments-body-focused .see-docx-comment-bubble.active {{
  box-shadow: 0 0 0 1px {palette["highlight"]};
}}
.see-docx-comments.comments-body-focused .see-docx-comment-thread.active {{
  box-shadow: 0 0 0 1px {palette["highlight"]};
}}
.see-docx-comments-count {{
  color: {palette["metadata"]};
  font-size: 0.76em;
  font-weight: 700;
}}
.see-docx-comments-key {{
  color: {palette["muted"]};
  font-size: 0.68em;
  font-weight: 700;
  letter-spacing: 0.02em;
}}
.see-docx-comments.comments-focused .see-docx-comments-key {{
  color: {palette["highlight"]};
}}
.see-docx-comments-hint {{
  color: {palette["muted"]};
  font-size: 0.72em;
}}
.see-docx-comments-empty {{
  color: {palette["muted"]};
  font-size: 0.90em;
  padding: 18px {COMMENTS_PANEL_GUTTER}px;
}}
.see-docx-comments-summary {{
  color: {palette["muted"]};
  font-size: 0.68em;
}}
.see-docx-comment-thread {{
  background-color: {palette["panel"]};
  border: 1px solid {palette["separator"]};
  border-left: 2px solid {palette["separator"]};
  border-radius: 5px;
  padding: 9px 10px 10px 9px;
}}
.see-docx-comment-thread.active {{
  background-color: {palette["panel"]};
  border-left: 3px solid {palette["accent"]};
}}
.see-docx-comment-thread.active .see-docx-comment-body {{
  font-size: 1.00em;
}}
.see-docx-comment-thread.active .see-docx-comment-author {{
  font-size: 0.90em;
}}
.see-docx-comment-thread.active .see-docx-comment-anchor {{
  font-size: 0.82em;
}}
.see-docx-comment-thread-label {{
  color: {palette["muted"]};
  font-size: 0.68em;
  font-weight: 700;
  letter-spacing: 0.08em;
}}
.see-docx-comment-message.reply {{
  border-top: 1px solid {palette["separator"]};
  margin-left: 13px;
  padding-top: 8px;
}}
.see-docx-comment-reply-label {{
  color: {palette["accent"]};
  font-size: 0.68em;
  font-weight: 700;
  letter-spacing: 0.06em;
}}
.see-docx-comment-bubble {{
  background-color: {palette["panel"]};
  border: 1px solid {palette["separator"]};
  border-left: 2px solid {palette["separator"]};
  border-radius: 5px;
  padding: 10px 11px 11px 10px;
}}
.see-docx-comment-bubble.active {{
  background-color: {palette["panel"]};
  border-left: 3px solid {palette["accent"]};
}}
.see-docx-comment-bubble.active .see-docx-comment-body {{
  font-size: 1.00em;
}}
.see-docx-comment-bubble.active .see-docx-comment-author {{
  font-size: 0.90em;
}}
.see-docx-comment-bubble.active .see-docx-comment-anchor {{
  font-size: 0.82em;
}}
.see-docx-comment-body-scroll,
.see-docx-comment-body-scroll viewport {{
  background-color: {palette["panel"]};
  border: 0;
}}
.see-docx-comment-body-scroll scrollbar.vertical slider {{
  min-width: 4px;
  border-radius: 3px;
  background-color: {palette["separator"]};
}}
.see-docx-comment-initials {{
  background-color: {palette["accent"]};
  border-radius: 10px;
  color: {palette["background"]};
  font-size: 0.72em;
  font-weight: 800;
  min-width: 20px;
  padding: 3px 4px;
}}
.see-docx-comment-author {{
  color: {palette["highlight"]};
  font-size: 0.86em;
  font-weight: 700;
}}
.see-docx-comment-date {{
  color: {palette["muted"]};
  font-size: 0.75em;
}}
.see-docx-comment-body {{
  color: {palette["foreground"]};
  font-size: 0.90em;
}}
.see-docx-comment-anchor {{
  color: {palette["muted"]};
  font-size: 0.78em;
  font-style: italic;
}}
.see-docx-search {{
  background-color: {palette["panel"]};
  border: 1px solid {palette["separator"]};
  box-shadow: 0 4px 14px rgba(0, 0, 0, 0.28);
  min-height: 38px;
  padding: 0 10px;
}}
.see-docx-search entry {{
  min-width: 0;
  background-color: transparent;
  border: 0;
  box-shadow: none;
  color: {palette["view_foreground"]};
  padding: 8px 6px;
}}
.see-docx-search-prefix {{
  color: {palette["highlight"]};
  font-weight: 700;
  padding: 0 8px 0 2px;
}}
.see-docx-search-status {{
  color: {palette["muted"]};
  font-size: 0.85em;
  padding: 0 2px 0 8px;
}}
.see-docx-page-jump {{
  background-color: {palette["panel"]};
  border: 1px solid {palette["separator"]};
  box-shadow: 0 4px 14px rgba(0, 0, 0, 0.28);
  padding: 5px 8px;
}}
.see-docx-page-jump-prefix {{
  color: {palette["highlight"]};
  font-weight: 700;
  padding: 0 5px 0 2px;
}}
.see-docx-page-jump entry {{
  min-width: 260px;
  background-color: {palette["panel_dark"]};
  color: {palette["view_foreground"]};
}}
.see-docx-page-jump-status {{
  color: {palette["muted"]};
  font-size: 0.85em;
  min-width: 48px;
  padding-left: 8px;
}}
.see-docx-export {{
  background-color: {palette["panel"]};
  border: 1px solid {palette["separator"]};
  box-shadow: 0 4px 14px rgba(0, 0, 0, 0.28);
  min-width: 300px;
  padding: 10px;
}}
.see-docx-export-title {{
  color: {palette["muted"]};
  font-size: 0.80em;
  font-weight: 700;
  letter-spacing: 0.08em;
  padding: 2px 4px 7px;
}}
.see-docx-export list {{
  background-color: transparent;
  color: {palette["view_foreground"]};
}}
.see-docx-export row {{
  border-radius: 3px;
  padding: 6px 8px;
}}
.see-docx-export row:selected {{
  background-color: {palette["selected_background"]};
  color: {palette["selected_foreground"]};
}}
.see-docx-export-format {{
  font-weight: 700;
}}
.see-docx-export-description,
.see-docx-export-status {{
  color: {palette["muted"]};
  font-size: 0.85em;
}}
.see-docx-export-status {{
  padding: 8px 4px 2px;
}}
.see-docx-status {{
  background-color: {palette["panel_dark"]};
  padding: 6px 12px;
}}
.see-docx-search-session {{
  color: {palette["highlight"]};
  font-size: 0.78em;
  font-weight: 700;
}}
.see-docx-reading-progress,
.see-docx-reading-progress trough {{
  min-height: 3px;
  background-color: {palette["separator"]};
  border: 0;
  border-radius: 0;
}}
.see-docx-reading-progress progress {{
  min-height: 3px;
  background-color: {palette["accent"]};
  border: 0;
  border-radius: 0;
}}
.see-docx-path-status {{
  color: {palette["metadata"]};
  font-size: 0.82em;
}}
.see-docx-url-hint {{
  background-color: rgba(2, 3, 3, 0.60);
  border: 1px solid {palette["accent"]};
  border-radius: 4px;
  color: {palette["accent"]};
  font-family: "Eurostile Next Pro", "Eurotype";
  font-size: 10pt;
  padding: 2px 6px;
}}
""".encode()


@dataclass(frozen=True)
class OutlineEntry:
    """A navigable heading destination from the PDF document outline."""

    title: str
    page_index: int
    top: float | None
    depth: int


@dataclass(frozen=True)
class OutlineLocatorFrame:
    """The visual state of the outline destination marker at one instant."""

    expansion: float
    bloom_opacity: float
    fill_opacity: float
    anchor_opacity: float


def _clamp_unit(value: float) -> float:
    return min(max(value, 0.0), 1.0)


def _reading_progress_fraction(scroll: float, maximum_scroll: float) -> float:
    """Map the document's vertical offset to a passive reading indicator."""

    if maximum_scroll <= 0.0:
        return 0.0
    return _clamp_unit(scroll / maximum_scroll)


def _ease_out_cubic(value: float) -> float:
    """Ease a unit interval quickly into a calm, settled finish."""

    return 1.0 - (1.0 - _clamp_unit(value)) ** 3


def _reading_progress_frame(start: float, target: float, elapsed_ms: float) -> float:
    """Return one eased frame for the passive reading-progress rule."""

    progress = _ease_out_cubic(elapsed_ms / READING_PROGRESS_DURATION_MS)
    return _clamp_unit(start + (target - start) * progress)


def _ease_in_out_cubic(value: float) -> float:
    """Give a visible sweep a gentle start and a settled finish."""

    value = _clamp_unit(value)
    if value < 0.5:
        return 4.0 * value**3
    return 1.0 - (-2.0 * value + 2.0) ** 3 / 2.0


def _outline_locator_frame(elapsed_ms: float) -> OutlineLocatorFrame:
    """Return the clipped-ripple motion and fade for an outline arrival."""

    elapsed_ms = max(elapsed_ms, 0.0)
    fade = 1.0 - _clamp_unit(
        (elapsed_ms - OUTLINE_LOCATOR_RIPPLE_FADE_START_MS)
        / (OUTLINE_LOCATOR_DURATION_MS - OUTLINE_LOCATOR_RIPPLE_FADE_START_MS)
    )
    return OutlineLocatorFrame(
        expansion=_ease_in_out_cubic(elapsed_ms / OUTLINE_LOCATOR_RIPPLE_EXPANSION_MS),
        bloom_opacity=(1.0 - _clamp_unit(elapsed_ms / OUTLINE_LOCATOR_BLOOM_MS)) * fade,
        fill_opacity=fade,
        anchor_opacity=(0.55 + 0.45 * _ease_out_cubic(elapsed_ms / 100.0)) * fade,
    )


def _outline_default_expansion_depth(entries: Iterable[OutlineEntry]) -> int:
    """Choose the deepest whole heading level that fits the initial outline."""

    total = 0
    depth = 0
    while True:
        count = sum(entry.depth == depth for entry in entries)
        if count == 0 or total + count > OUTLINE_INITIAL_VISIBLE_LIMIT:
            return max(0, depth - 1)
        total += count
        depth += 1


def _contextual_scroll_target(
    destination: float, viewport_height: float, maximum_scroll: float
) -> float:
    """Keep a heading in context by placing it below the viewport's top edge."""

    target = destination - viewport_height * OUTLINE_CONTEXT_VIEWPORT_FRACTION
    return min(max(target, 0.0), maximum_scroll)


def _search_scroll_target(
    destination: float, viewport_height: float, maximum_scroll: float
) -> float:
    """Center a text-search result in the viewport whenever space permits."""

    target = destination - viewport_height * SEARCH_CONTEXT_VIEWPORT_FRACTION
    return min(max(target, 0.0), maximum_scroll)


def _outline_header_line_height(page: Poppler.Page, top: float) -> float:
    """Return the rendered height of the text line at an outline destination."""

    try:
        has_layout, rectangles = page.get_text_layout()
    except (AttributeError, TypeError):
        return OUTLINE_LOCATOR_FALLBACK_HEIGHT
    if not has_layout:
        return OUTLINE_LOCATOR_FALLBACK_HEIGHT

    _page_width, page_height = page.get_size()
    # Outline destinations use the PDF's bottom-origin coordinates; Poppler
    # text layout rectangles use the rendered page's top-origin coordinates.
    destination_y = page_height - top
    glyphs = [
        rectangle
        for rectangle in rectangles
        if float(rectangle.y2) > float(rectangle.y1)
    ]
    if not glyphs:
        return OUTLINE_LOCATOR_FALLBACK_HEIGHT

    def line_center(rectangle: Poppler.Rectangle) -> float:
        return (float(rectangle.y1) + float(rectangle.y2)) / 2.0

    nearest = min(
        glyphs, key=lambda rectangle: abs(line_center(rectangle) - destination_y)
    )
    nearest_center = line_center(nearest)
    # Glyph rectangles within a text line share a baseline.  Grouping by the
    # vertical centre lets the marker retain the full line height even where
    # a heading mixes glyph shapes with slightly different ascenders.
    line = [
        rectangle
        for rectangle in glyphs
        if abs(line_center(rectangle) - nearest_center) <= 1.0
    ]
    return max(
        1.0,
        max(float(rectangle.y2) for rectangle in line)
        - min(float(rectangle.y1) for rectangle in line),
    )


@dataclass(frozen=True)
class SearchMatch:
    """A Poppler text match and the document position it identifies."""

    page_index: int
    left: float
    bottom: float
    right: float
    top: float


@dataclass(frozen=True)
class UrlLink:
    """An external URI action and its rectangle in PDF coordinates."""

    uri: str
    left: float
    bottom: float
    right: float
    top: float


@dataclass(frozen=True)
class DocumentComment:
    """One Word comment and the source range it annotates."""

    comment_id: str
    author: str
    initials: str
    date: str | None
    text: str
    source_start: int
    source_end: int
    parent_comment_id: str | None = None


@dataclass(frozen=True)
class CommentThread:
    """One anchored conversation, with its root comment followed by replies."""

    thread_id: str
    comments: tuple[DocumentComment, ...]

    @property
    def root(self) -> DocumentComment:
        return self.comments[0]

    @property
    def reply_count(self) -> int:
        return max(0, len(self.comments) - 1)

    # These aliases keep the navigation model thread-shaped without making
    # callers care whether they are looking at a single comment or a thread.
    @property
    def comment_id(self) -> str:
        return self.thread_id

    @property
    def author(self) -> str:
        return self.root.author

    @property
    def initials(self) -> str:
        return self.root.initials

    @property
    def date(self) -> str | None:
        return self.root.date

    @property
    def text(self) -> str:
        return self.root.text


@dataclass(frozen=True)
class CommentAnchor:
    """Rendered glyph rectangles for one comment's first anchored page."""

    comment: DocumentComment
    page_index: int
    rectangles: tuple[tuple[float, float, float, float], ...]

    @property
    def right(self) -> float:
        return max(rectangle[2] for rectangle in self.rectangles)

    @property
    def center_y(self) -> float:
        return sum(rectangle[1] + rectangle[3] for rectangle in self.rectangles) / (
            2 * len(self.rectangles)
        )


@dataclass(frozen=True)
class CommentMark:
    """One rendered glyph rectangle associated with a DOCX comment."""

    comment_id: str
    rectangle: tuple[float, float, float, float]


@dataclass(frozen=True)
class TextSelection:
    """A normalized, top-origin rectangular selection on one PDF page."""

    left: float
    top: float
    right: float
    bottom: float


@dataclass(frozen=True)
class RichTextClipboardPayload:
    """The interoperable rich and plain representations of a selection."""

    text: str
    html: str


@dataclass(frozen=True)
class _DocxRun:
    """One DOCX run, retaining the inline properties Writer needs on paste."""

    text: str
    tags: tuple[tuple[str, str], ...]

    def html_for_range(self, start: int, end: int) -> str:
        """Render the requested local character range with this run's style."""

        text = self.text[max(start, 0) : max(end, 0)]
        if not text:
            return ""
        rendered = html_escape(text, quote=False).replace("\n", "<br>")
        for opening, closing in reversed(self.tags):
            rendered = f"{opening}{rendered}{closing}"
        return rendered


@dataclass(frozen=True)
class _DocxParagraph:
    """A paragraph whose runs can be copied without losing inline style."""

    runs: tuple[_DocxRun, ...]
    opening_tag: str = "<p>"

    @property
    def text(self) -> str:
        return "".join(run.text for run in self.runs)

    def html_for_range(self, start: int, end: int) -> str:
        """Render a local selection as one valid HTML paragraph."""

        cursor = 0
        content: list[str] = []
        for run in self.runs:
            run_end = cursor + len(run.text)
            overlap_start = max(start, cursor)
            overlap_end = min(end, run_end)
            if overlap_start < overlap_end:
                content.append(
                    run.html_for_range(overlap_start - cursor, overlap_end - cursor)
                )
            cursor = run_end
        return f"{self.opening_tag}{''.join(content)}</p>"


@dataclass(frozen=True)
class _DocxTableCell:
    """A DOCX table cell, including its own rich paragraph content."""

    paragraphs: tuple[_DocxParagraph, ...]
    colspan: int = 1

    @property
    def text(self) -> str:
        return "\n".join(paragraph.text for paragraph in self.paragraphs)

    def html_for_range(self, start: int, end: int) -> str:
        """Copy a cell as a one-cell table, matching Writer's table semantics."""

        cursor = 0
        paragraphs: list[str] = []
        for paragraph in self.paragraphs:
            paragraph_end = cursor + len(paragraph.text)
            overlap_start = max(start, cursor)
            overlap_end = min(end, paragraph_end)
            if overlap_start < overlap_end:
                paragraphs.append(
                    paragraph.html_for_range(
                        overlap_start - cursor, overlap_end - cursor
                    )
                )
            cursor = paragraph_end + 1
        colspan = f' colspan="{self.colspan}"' if self.colspan > 1 else ""
        return f"<table><tbody><tr><td{colspan}>{''.join(paragraphs)}</td></tr></tbody></table>"

    def full_html(self) -> str:
        return self.html_for_range(0, len(self.text))


@dataclass(frozen=True)
class _DocxTable:
    """A table in the DOCX body, retained for copy-all rich text."""

    rows: tuple[tuple[_DocxTableCell, ...], ...]

    @property
    def selection_text(self) -> str:
        """Return the glyph order emitted by Writer's PDF exporter."""

        return "\n".join("".join(cell.text for cell in row) for row in self.rows)

    @property
    def plain_text(self) -> str:
        return "\n".join("\t".join(cell.text for cell in row) for row in self.rows)

    def full_html(self) -> str:
        rows: list[str] = []
        for row in self.rows:
            cells = []
            for cell in row:
                colspan = f' colspan="{cell.colspan}"' if cell.colspan > 1 else ""
                content = "".join(
                    paragraph.html_for_range(0, len(paragraph.text))
                    for paragraph in cell.paragraphs
                )
                cells.append(f"<td{colspan}>{content}</td>")
            rows.append(f"<tr>{''.join(cells)}</tr>")
        return f"<table><tbody>{''.join(rows)}</tbody></table>"


@dataclass(frozen=True)
class _DocxSelectionFragment:
    """A selectable paragraph or cell positioned in PDF glyph order."""

    start: int
    end: int
    content: _DocxParagraph | _DocxTableCell

    @property
    def text(self) -> str:
        return self.content.text

    def html_for_range(self, start: int, end: int) -> str:
        return self.content.html_for_range(start, end)


class _DocxRichTextSource:
    """DOCX source structure used to restore rich selections from PDF glyphs."""

    def __init__(
        self,
        blocks: tuple[_DocxParagraph | _DocxTable, ...],
        fragments: tuple[_DocxSelectionFragment, ...],
        selection_text: str,
        comments: tuple[DocumentComment, ...] = (),
    ) -> None:
        self._blocks = blocks
        self._fragments = fragments
        self._selection_text = selection_text
        self.comments = comments

    @property
    def plain_text(self) -> str:
        """Return the document text with table columns preserved by tabs."""

        return "\n".join(
            block.text if isinstance(block, _DocxParagraph) else block.plain_text
            for block in self._blocks
        )

    def full_payload(self) -> RichTextClipboardPayload:
        """Return a rich representation of the complete source document."""

        content = "".join(
            block.html_for_range(0, len(block.text))
            if isinstance(block, _DocxParagraph)
            else block.full_html()
            for block in self._blocks
        )
        return RichTextClipboardPayload(
            text=self.plain_text,
            html=f"<html><body>{content}</body></html>",
        )

    def payload_for_range(
        self, start: int, end: int
    ) -> RichTextClipboardPayload | None:
        """Render an unambiguous source range as rich clipboard data."""

        start = max(start, 0)
        end = min(max(end, start), len(self._selection_text))
        if start == end:
            return None
        fragments = [
            fragment
            for fragment in self._fragments
            if fragment.start < end and start < fragment.end
        ]
        if not fragments:
            return None
        content: list[str] = []
        for fragment in fragments:
            content.append(
                fragment.html_for_range(
                    max(start, fragment.start) - fragment.start,
                    min(end, fragment.end) - fragment.start,
                )
            )
        return RichTextClipboardPayload(
            text=self._selection_text[start:end],
            html=f"<html><body>{''.join(content)}</body></html>",
        )

    def payload_for_text(self, selected_text: str) -> RichTextClipboardPayload | None:
        """Find one selected PDF string in the source, refusing ambiguity."""

        selected, _selected_indices = _fold_selection_whitespace(selected_text)
        source, source_indices = _fold_selection_whitespace(self._selection_text)
        if not selected:
            return None
        start = source.find(selected)
        if start < 0 or source.find(selected, start + 1) >= 0:
            return None
        end = start + len(selected) - 1
        return self.payload_for_range(source_indices[start], source_indices[end] + 1)


class _GtkTargetEntry(ctypes.Structure):
    """The C structure used by GTK3's multi-target clipboard API."""

    _fields_ = [
        ("target", ctypes.c_char_p),
        ("flags", ctypes.c_uint),
        ("info", ctypes.c_uint),
    ]


_GtkClipboardGetFunc = ctypes.CFUNCTYPE(
    None,
    ctypes.c_void_p,
    ctypes.c_void_p,
    ctypes.c_uint,
    ctypes.c_void_p,
)
_GtkClipboardClearFunc = ctypes.CFUNCTYPE(None, ctypes.c_void_p, ctypes.c_void_p)
_RICH_CLIPBOARD_HTML_INFO = 1
_RICH_CLIPBOARD_TEXT_INFO = 2
_RICH_CLIPBOARD_TARGETS = (
    (b"text/html", _RICH_CLIPBOARD_HTML_INFO),
    (b"text/plain;charset=utf-8", _RICH_CLIPBOARD_TEXT_INFO),
    (b"UTF8_STRING", _RICH_CLIPBOARD_TEXT_INFO),
    (b"TEXT", _RICH_CLIPBOARD_TEXT_INFO),
    (b"STRING", _RICH_CLIPBOARD_TEXT_INFO),
)
_native_gtk_clipboard: ctypes.CDLL | None = None
_native_gtk_clipboard_loaded = False
_native_clipboard_owners: dict[int, RichTextClipboardPayload] = {}


def _native_gtk_clipboard_api() -> ctypes.CDLL | None:
    """Load GTK3's callback clipboard API omitted from PyGObject bindings."""

    global _native_gtk_clipboard, _native_gtk_clipboard_loaded
    if _native_gtk_clipboard_loaded:
        return _native_gtk_clipboard
    _native_gtk_clipboard_loaded = True
    try:
        library = ctypes.CDLL(ctypes.util.find_library("gtk-3") or "libgtk-3.so.0")
    except OSError:
        return None
    library.gtk_clipboard_set_with_data.argtypes = [
        ctypes.c_void_p,
        ctypes.POINTER(_GtkTargetEntry),
        ctypes.c_uint,
        _GtkClipboardGetFunc,
        _GtkClipboardClearFunc,
        ctypes.c_void_p,
    ]
    library.gtk_clipboard_set_with_data.restype = ctypes.c_int
    library.gtk_selection_data_get_target.argtypes = [ctypes.c_void_p]
    library.gtk_selection_data_get_target.restype = ctypes.c_void_p
    library.gtk_selection_data_set.argtypes = [
        ctypes.c_void_p,
        ctypes.c_void_p,
        ctypes.c_int,
        ctypes.POINTER(ctypes.c_ubyte),
        ctypes.c_int,
    ]
    library.gtk_selection_data_set.restype = None
    library.gtk_selection_data_set_text.argtypes = [
        ctypes.c_void_p,
        ctypes.c_char_p,
        ctypes.c_int,
    ]
    library.gtk_selection_data_set_text.restype = None
    _native_gtk_clipboard = library
    return library


@_GtkClipboardGetFunc
def _on_native_clipboard_get(
    clipboard: int, selection_data: int, info: int, _user_data: int
) -> None:
    """Answer a GTK clipboard request with the requested MIME representation."""

    payload = _native_clipboard_owners.get(int(clipboard))
    library = _native_gtk_clipboard
    if payload is None or library is None:
        return
    if info == _RICH_CLIPBOARD_HTML_INFO:
        data = payload.html.encode("utf-8")
        buffer = (ctypes.c_ubyte * len(data)).from_buffer_copy(data)
        target = library.gtk_selection_data_get_target(selection_data)
        library.gtk_selection_data_set(selection_data, target, 8, buffer, len(data))
        return
    data = payload.text.encode("utf-8")
    library.gtk_selection_data_set_text(selection_data, data, len(data))


@_GtkClipboardClearFunc
def _on_native_clipboard_clear(clipboard: int, _user_data: int) -> None:
    """Release the Python payload after GTK no longer owns the selection."""

    _native_clipboard_owners.pop(int(clipboard), None)


def _publish_rich_clipboard(
    clipboard: Gtk.Clipboard, payload: RichTextClipboardPayload
) -> bool:
    """Offer HTML and plain text through GTK3's native selection owner."""

    library = _native_gtk_clipboard_api()
    if library is None or not hasattr(clipboard, "__gpointer__"):
        return False
    entries = (_GtkTargetEntry * len(_RICH_CLIPBOARD_TARGETS))(
        *(
            _GtkTargetEntry(target, 0, info)
            for target, info in _RICH_CLIPBOARD_TARGETS
        )
    )
    address = hash(clipboard)
    published = library.gtk_clipboard_set_with_data(
        ctypes.c_void_p(address),
        entries,
        len(entries),
        _on_native_clipboard_get,
        _on_native_clipboard_clear,
        None,
    )
    if not published:
        return False
    _native_clipboard_owners[address] = payload
    return True


def _reflow_pdf_selection_text(text: str) -> str:
    """Join PDF-rendered line wraps so pasted prose remains reflowable."""

    lines = [line.strip() for line in text.splitlines() if line.strip()]
    if not lines:
        return ""
    reflowed = lines[0]
    for line in lines[1:]:
        # Preserve an explicit visible hyphen when a word breaks at a line
        # edge, while treating ordinary visual line ends as whitespace.
        separator = "" if reflowed.endswith(("-", "‐", "\u00ad")) else " "
        reflowed = f"{reflowed}{separator}{line}"
    return reflowed


def _fold_selection_whitespace(text: str) -> tuple[str, list[int]]:
    """Collapse whitespace while retaining each folded character's source index."""

    folded: list[str] = []
    indices: list[int] = []
    for index, character in enumerate(text):
        if character.isspace():
            if folded and folded[-1] != " ":
                folded.append(" ")
                indices.append(index)
        else:
            folded.append(character)
            indices.append(index)
    if folded and folded[-1] == " ":
        folded.pop()
        indices.pop()
    return "".join(folded), indices


def _restore_docx_paragraph_boundaries(
    source_text: str, selected_text: str
) -> str | None:
    """Recover source paragraph breaks obscured by PDF visual line wrapping."""

    source, source_indices = _fold_selection_whitespace(source_text)
    selected, _selected_indices = _fold_selection_whitespace(selected_text)
    if not selected:
        return ""
    start = source.find(selected)
    if start < 0 or source.find(selected, start + 1) >= 0:
        return None
    end = start + len(selected) - 1
    return source_text[source_indices[start] : source_indices[end] + 1]


def _docx_property_enabled(element: ET.Element | None) -> bool:
    """Return whether a WordprocessingML boolean property is enabled."""

    if element is None:
        return False
    value = element.get(f"{_WORDPROCESSINGML}val", "true").lower()
    return value not in {"0", "false", "off", "none"}


def _docx_run_tags(run: ET.Element, href: str | None = None) -> tuple[tuple[str, str], ...]:
    """Translate portable DOCX run properties into an HTML tag stack."""

    properties = run.find(f"{_WORDPROCESSINGML}rPr")
    style: list[str] = []
    if properties is not None:
        color = properties.find(f"{_WORDPROCESSINGML}color")
        color_value = color.get(f"{_WORDPROCESSINGML}val") if color is not None else None
        if color_value and len(color_value) == 6 and all(
            character in "0123456789abcdefABCDEF" for character in color_value
        ):
            style.append(f"color: #{color_value}")
        size = properties.find(f"{_WORDPROCESSINGML}sz")
        size_value = size.get(f"{_WORDPROCESSINGML}val") if size is not None else None
        if size_value and size_value.isdigit():
            style.append(f"font-size: {int(size_value) / 2:g}pt")
        fonts = properties.find(f"{_WORDPROCESSINGML}rFonts")
        font_name = None
        if fonts is not None:
            font_name = fonts.get(f"{_WORDPROCESSINGML}ascii") or fonts.get(
                f"{_WORDPROCESSINGML}hAnsi"
            )
        if font_name:
            style.append(f"font-family: {html_escape(font_name, quote=True)}")

    tags: list[tuple[str, str]] = []
    if href:
        tags.append((f'<a href="{html_escape(href, quote=True)}">', "</a>"))
    if style:
        tags.append((f'<span style="{"; ".join(style)}">', "</span>"))
    if _docx_property_enabled(
        properties.find(f"{_WORDPROCESSINGML}b") if properties is not None else None
    ):
        tags.append(("<strong>", "</strong>"))
    if _docx_property_enabled(
        properties.find(f"{_WORDPROCESSINGML}i") if properties is not None else None
    ):
        tags.append(("<em>", "</em>"))
    underline = properties.find(f"{_WORDPROCESSINGML}u") if properties is not None else None
    if _docx_property_enabled(underline):
        tags.append(("<u>", "</u>"))
    if _docx_property_enabled(
        properties.find(f"{_WORDPROCESSINGML}strike") if properties is not None else None
    ):
        tags.append(("<s>", "</s>"))
    vertical = (
        properties.find(f"{_WORDPROCESSINGML}vertAlign")
        if properties is not None
        else None
    )
    if vertical is not None:
        value = vertical.get(f"{_WORDPROCESSINGML}val")
        if value == "superscript":
            tags.append(("<sup>", "</sup>"))
        elif value == "subscript":
            tags.append(("<sub>", "</sub>"))
    return tuple(tags)


def _docx_run(run: ET.Element, href: str | None = None) -> _DocxRun | None:
    """Extract one visible Word run and the formatting used to render it."""

    text: list[str] = []
    for element in run:
        if element.tag == f"{_WORDPROCESSINGML}t":
            text.append(element.text or "")
        elif element.tag == f"{_WORDPROCESSINGML}tab":
            text.append("\t")
        elif element.tag in {
            f"{_WORDPROCESSINGML}br",
            f"{_WORDPROCESSINGML}cr",
        }:
            text.append("\n")
        elif element.tag == f"{_WORDPROCESSINGML}noBreakHyphen":
            text.append("‑")
    content = "".join(text)
    return _DocxRun(content, _docx_run_tags(run, href)) if content else None


def _docx_paragraph(
    element: ET.Element, relationships: dict[str, str]
) -> _DocxParagraph:
    """Parse a paragraph's visible runs while retaining their inline style."""

    properties = element.find(f"{_WORDPROCESSINGML}pPr")
    opening_tag = "<p>"
    if properties is not None:
        alignment = properties.find(f"{_WORDPROCESSINGML}jc")
        value = alignment.get(f"{_WORDPROCESSINGML}val") if alignment is not None else None
        alignments = {"left", "center", "right", "both", "justify"}
        if value in alignments:
            css_value = "justify" if value == "both" else value
            opening_tag = f'<p style="text-align: {css_value}">'

    runs: list[_DocxRun] = []
    for child in element:
        if child.tag == f"{_WORDPROCESSINGML}r":
            run = _docx_run(child)
            if run is not None:
                runs.append(run)
        elif child.tag == f"{_WORDPROCESSINGML}hyperlink":
            relationship_id = child.get(f"{_WORDPROCESSINGML}id")
            href = relationships.get(relationship_id or "")
            for run_element in child.findall(f"{_WORDPROCESSINGML}r"):
                run = _docx_run(run_element, href)
                if run is not None:
                    runs.append(run)
    return _DocxParagraph(tuple(runs), opening_tag)


def _docx_table(
    element: ET.Element, relationships: dict[str, str]
) -> _DocxTable:
    """Parse the table grid needed for table-aware copy and paste."""

    rows: list[tuple[_DocxTableCell, ...]] = []
    for row_element in element.findall(f"{_WORDPROCESSINGML}tr"):
        cells: list[_DocxTableCell] = []
        for cell_element in row_element.findall(f"{_WORDPROCESSINGML}tc"):
            paragraphs = tuple(
                _docx_paragraph(paragraph, relationships)
                for paragraph in cell_element.findall(f"{_WORDPROCESSINGML}p")
            )
            cell_properties = cell_element.find(f"{_WORDPROCESSINGML}tcPr")
            grid_span = (
                cell_properties.find(f"{_WORDPROCESSINGML}gridSpan")
                if cell_properties is not None
                else None
            )
            span_value = grid_span.get(f"{_WORDPROCESSINGML}val") if grid_span is not None else None
            colspan = int(span_value) if span_value and span_value.isdigit() else 1
            cells.append(_DocxTableCell(paragraphs, colspan))
        rows.append(tuple(cells))
    return _DocxTable(tuple(rows))


def _docx_relationships(archive: zipfile.ZipFile) -> dict[str, str]:
    """Return external hyperlink targets keyed by the document relationship ID."""

    try:
        relationships = ET.fromstring(archive.read("word/_rels/document.xml.rels"))
    except (ET.ParseError, KeyError):
        return {}
    result: dict[str, str] = {}
    for relationship in relationships:
        if relationship.get("TargetMode") != "External":
            continue
        identifier = relationship.get("Id")
        target = relationship.get("Target")
        if identifier and target:
            result[identifier] = target
    return result


def _docx_element_text(element: ET.Element) -> str:
    """Extract visible Word text from a comment body or inline container."""

    text: list[str] = []
    for child in element.iter():
        if child.tag == f"{_WORDPROCESSINGML}t":
            text.append(child.text or "")
        elif child.tag == f"{_WORDPROCESSINGML}tab":
            text.append("\t")
        elif child.tag in {
            f"{_WORDPROCESSINGML}br",
            f"{_WORDPROCESSINGML}cr",
        }:
            text.append("\n")
    return "".join(text)


def _docx_comment_records(
    archive: zipfile.ZipFile,
) -> dict[str, tuple[str, str, str | None, str, str | None]]:
    """Read comment metadata without making the comments part mandatory."""

    try:
        comments = ET.fromstring(archive.read("word/comments.xml"))
    except (ET.ParseError, KeyError):
        return {}
    paragraph_to_comment: dict[str, str] = {}
    for comment in comments.findall(f"{_WORDPROCESSINGML}comment"):
        identifier = comment.get(f"{_WORDPROCESSINGML}id")
        if identifier is None:
            continue
        for paragraph in comment.findall(f"{_WORDPROCESSINGML}p"):
            paragraph_id = paragraph.get(f"{_WORD_2010_WORDML}paraId")
            if paragraph_id:
                paragraph_to_comment[paragraph_id] = identifier

    parent_by_comment: dict[str, str | None] = {}
    try:
        comments_extended = ET.fromstring(
            archive.read("word/commentsExtended.xml")
        )
    except (ET.ParseError, KeyError):
        comments_extended = None
    if comments_extended is not None:
        for comment_ex in comments_extended.findall(
            f"{_WORD_2012_WORDML}commentEx"
        ):
            paragraph_id = comment_ex.get(f"{_WORD_2012_WORDML}paraId")
            identifier = paragraph_to_comment.get(paragraph_id or "")
            if identifier is None:
                continue
            parent_paragraph_id = comment_ex.get(
                f"{_WORD_2012_WORDML}paraIdParent"
            )
            parent_by_comment[identifier] = (
                paragraph_to_comment.get(parent_paragraph_id)
                if parent_paragraph_id
                else None
            )

    records: dict[str, tuple[str, str, str | None, str, str | None]] = {}
    for comment in comments.findall(f"{_WORDPROCESSINGML}comment"):
        identifier = comment.get(f"{_WORDPROCESSINGML}id")
        if identifier is None:
            continue
        author = comment.get(f"{_WORDPROCESSINGML}author") or "Comment"
        initials = comment.get(f"{_WORDPROCESSINGML}initials") or ""
        date = comment.get(f"{_WORDPROCESSINGML}date")
        paragraphs = [
            _docx_element_text(paragraph).strip()
            for paragraph in comment.findall(f"{_WORDPROCESSINGML}p")
        ]
        text = "\n".join(paragraph for paragraph in paragraphs if paragraph)
        if not initials:
            initials = "".join(
                part[0] for part in author.split() if part
            ).upper()[:2]
        records[identifier] = (
            author,
            initials,
            date,
            text,
            parent_by_comment.get(identifier),
        )
    return records


def _comment_threads(
    comments: Iterable[DocumentComment],
) -> tuple[CommentThread, ...]:
    """Group root comments and replies while preserving document order."""

    ordered = tuple(comments)
    by_id = {comment.comment_id: comment for comment in ordered}
    order = {comment.comment_id: index for index, comment in enumerate(ordered)}
    grouped: dict[str, list[DocumentComment]] = {}

    def root_id(identifier: str) -> str:
        current = identifier
        visited: set[str] = set()
        while current not in visited:
            visited.add(current)
            comment = by_id.get(current)
            parent = comment.parent_comment_id if comment is not None else None
            if parent is None or parent not in by_id:
                return current
            current = parent
        return identifier

    for comment in ordered:
        grouped.setdefault(root_id(comment.comment_id), []).append(comment)

    threads: list[CommentThread] = []
    for thread_id, members in sorted(
        grouped.items(), key=lambda item: order.get(item[0], 0)
    ):
        members.sort(
            key=lambda comment: (
                0 if comment.comment_id == thread_id else 1,
                order.get(comment.comment_id, 0),
            )
        )
        threads.append(CommentThread(thread_id, tuple(members)))
    return tuple(threads)


def _docx_comment_ranges(
    element: ET.Element, base_offset: int
) -> dict[str, tuple[int, int]]:
    """Return comment source ranges while counting only visible run text."""

    starts: dict[str, int] = {}
    ranges: dict[str, tuple[int, int]] = {}
    offset = base_offset

    def visit(node: ET.Element) -> None:
        nonlocal offset
        if node.tag == f"{_WORDPROCESSINGML}commentRangeStart":
            identifier = node.get(f"{_WORDPROCESSINGML}id")
            if identifier is not None:
                starts[identifier] = offset
            return
        if node.tag == f"{_WORDPROCESSINGML}commentRangeEnd":
            identifier = node.get(f"{_WORDPROCESSINGML}id")
            if identifier is not None and identifier in starts:
                ranges[identifier] = (starts[identifier], offset)
            return
        if node.tag == f"{_WORDPROCESSINGML}r":
            run = _docx_run(node)
            if run is not None:
                offset += len(run.text)
            return
        for child in node:
            visit(child)

    for child in element:
        visit(child)
    return ranges


def _docx_rich_text_source(path: Path) -> _DocxRichTextSource | None:
    """Parse enough OOXML structure to restore rich clipboard selections."""

    try:
        with zipfile.ZipFile(path) as archive:
            document = ET.fromstring(archive.read("word/document.xml"))
            relationships = _docx_relationships(archive)
            comment_records = _docx_comment_records(archive)
    except (ET.ParseError, KeyError, OSError, zipfile.BadZipFile):
        return None

    body = document.find(f"{_WORDPROCESSINGML}body")
    if body is None:
        return None
    blocks: list[_DocxParagraph | _DocxTable] = []
    fragments: list[_DocxSelectionFragment] = []
    selection_parts: list[str] = []
    comment_ranges: dict[str, tuple[int, int]] = {}
    cursor = 0

    def append(text: str) -> tuple[int, int]:
        nonlocal cursor
        start = cursor
        selection_parts.append(text)
        cursor += len(text)
        return start, cursor

    for element in body:
        if element.tag == f"{_WORDPROCESSINGML}p":
            paragraph = _docx_paragraph(element, relationships)
            blocks.append(paragraph)
            comment_ranges.update(_docx_comment_ranges(element, cursor))
            start, end = append(paragraph.text)
            if start != end:
                fragments.append(_DocxSelectionFragment(start, end, paragraph))
            append("\n")
        elif element.tag == f"{_WORDPROCESSINGML}tbl":
            table = _docx_table(element, relationships)
            blocks.append(table)
            for row_element, row in zip(
                element.findall(f"{_WORDPROCESSINGML}tr"),
                table.rows,
                strict=True,
            ):
                for cell_element, cell in zip(
                    row_element.findall(f"{_WORDPROCESSINGML}tc"),
                    row,
                    strict=True,
                ):
                    cell_cursor = cursor
                    for paragraph_index, (
                        paragraph_element,
                        paragraph,
                    ) in enumerate(
                        zip(
                            cell_element.findall(f"{_WORDPROCESSINGML}p"),
                            cell.paragraphs,
                            strict=True,
                        )
                    ):
                        comment_ranges.update(
                            _docx_comment_ranges(paragraph_element, cell_cursor)
                        )
                        cell_cursor += len(paragraph.text)
                        if paragraph_index + 1 < len(cell.paragraphs):
                            cell_cursor += 1
                    start, end = append(cell.text)
                    if start != end:
                        fragments.append(_DocxSelectionFragment(start, end, cell))
                append("\n")
    if not blocks:
        return None
    comments = tuple(
        DocumentComment(
            comment_id=identifier,
            author=comment_records[identifier][0],
            initials=comment_records[identifier][1],
            date=comment_records[identifier][2],
            text=comment_records[identifier][3],
            source_start=start,
            source_end=end,
            parent_comment_id=comment_records[identifier][4],
        )
        for identifier, (start, end) in sorted(
            comment_ranges.items(), key=lambda item: item[1]
        )
        if identifier in comment_records and start < end
    )
    return _DocxRichTextSource(
        tuple(blocks), tuple(fragments), "".join(selection_parts), comments
    )


def _docx_plain_text(path: Path) -> str | None:
    """Extract paragraph-preserving text from the stable DOCX source copy."""

    source = _docx_rich_text_source(path)
    return source.plain_text if source is not None else None


def _text_selection_bounds(
    *,
    start: tuple[float, float],
    end: tuple[float, float],
    page_width: float,
    page_height: float,
    line_padding: float,
) -> TextSelection:
    """Turn a drag into a page-clamped area Poppler can select glyphs from."""

    start_x, start_y = start
    end_x, end_y = end
    return TextSelection(
        left=min(max(min(start_x, end_x), 0.0), page_width),
        top=min(max(min(start_y, end_y) - line_padding, 0.0), page_height),
        right=min(max(max(start_x, end_x), 0.0), page_width),
        bottom=min(max(max(start_y, end_y) + line_padding, 0.0), page_height),
    )


def _document_search(document: Poppler.Document, query: str) -> list[SearchMatch]:
    """Find every literal text occurrence using Poppler's PDF text index."""

    if not query:
        return []
    matches: list[SearchMatch] = []
    for page_index in range(document.get_n_pages()):
        for rectangle in document.get_page(page_index).find_text(query):
            # PDF coordinates start at the bottom edge; ``destination_y``
            # expects the top edge in that same coordinate system.
            matches.append(
                SearchMatch(
                    page_index=page_index,
                    left=float(rectangle.x1),
                    bottom=float(rectangle.y1),
                    right=float(rectangle.x2),
                    top=float(rectangle.y2),
                )
            )
    return matches


def _page_url_links(page: Poppler.Page) -> list[UrlLink]:
    """Return externally openable URI actions from one rendered PDF page."""

    links: list[UrlLink] = []
    for mapping in page.get_link_mapping():
        action = mapping.action
        if action is None or action.type != Poppler.ActionType.URI:
            continue
        uri = action.uri.uri
        if not uri:
            continue
        area = mapping.area
        links.append(
            UrlLink(
                uri=uri,
                left=min(float(area.x1), float(area.x2)),
                bottom=min(float(area.y1), float(area.y2)),
                right=max(float(area.x1), float(area.x2)),
                top=max(float(area.y1), float(area.y2)),
            )
        )
    return links


def _outline_destination(
    document: Poppler.Document, destination: Poppler.Dest | None
) -> tuple[int, float | None] | None:
    """Turn a Poppler destination into a local page and optional PDF Y offset."""

    if destination is None:
        return None
    if destination.type == Poppler.DestType.NAMED:
        if not destination.named_dest:
            return None
        destination = document.find_dest(destination.named_dest)
        if destination is None:
            return None
    # Poppler uses the PDF's one-based page number, while the rendered pages
    # are stored in a zero-based Python list.
    page_index = destination.page_num - 1
    if not 0 <= page_index < document.get_n_pages():
        return None
    if destination.change_top:
        return page_index, float(destination.top)
    # LibreOffice exports outline destinations as XYZ bookmarks with a usable
    # Y coordinate but leaves ``change_top`` unset. For destination types
    # whose top coordinate has defined meaning, retain it when it falls on
    # the target page. Page-only destinations remain intentionally unmarked.
    if destination.type not in {
        Poppler.DestType.XYZ,
        Poppler.DestType.FITH,
        Poppler.DestType.FITBH,
    }:
        return page_index, None
    top = float(destination.top)
    _width, page_height = document.get_page(page_index).get_size()
    if 0 < top <= page_height:
        return page_index, top
    return page_index, None


def _outline_entries_from_iter(
    document: Poppler.Document, index: Poppler.IndexIter, depth: int = 0
) -> list[OutlineEntry]:
    """Flatten Poppler's outline tree while retaining its nesting depth."""

    entries: list[OutlineEntry] = []
    while True:
        action = index.get_action()
        if action.type == Poppler.ActionType.GOTO_DEST:
            destination = _outline_destination(document, action.goto_dest.dest)
            if destination is not None:
                page_index, top = destination
                title = action.goto_dest.title or "Untitled heading"
                entries.append(
                    OutlineEntry(
                        title=title,
                        page_index=page_index,
                        top=top,
                        depth=depth,
                    )
                )
        child = index.get_child()
        if child is not None:
            entries.extend(_outline_entries_from_iter(document, child, depth + 1))
        if not index.next():
            return entries


def _document_outline(document: Poppler.Document) -> list[OutlineEntry]:
    """Read the document's exported heading/bookmark hierarchy."""

    try:
        index = Poppler.IndexIter.new(document)
    except TypeError:
        # In current PyGObject bindings, the NULL iterator Poppler returns
        # for a PDF without bookmarks is surfaced as a TypeError instead.
        return []
    if index is None:
        return []
    # PyGObject releases the transfer-full iterator wrappers itself. Calling
    # their C ``free`` method explicitly leaves its finalizer with stale state.
    return _outline_entries_from_iter(document, index)


class PdfPage(Gtk.DrawingArea):
    """A lazily drawn Poppler page at the current zoom level."""

    def __init__(
        self,
        page: Poppler.Page,
        zoom: float,
        selection_handler: Callable[[str, "PdfPage", tuple[float, float]], None]
        | None = None,
        selection_scroll_handler: Callable[["PdfPage", Gdk.EventScroll], bool]
        | None = None,
    ) -> None:
        super().__init__()
        self._page = page
        self._zoom = zoom
        self._width, self._height = page.get_size()
        self._outline_locator_top: float | None = None
        self._outline_locator_height = OUTLINE_LOCATOR_FALLBACK_HEIGHT
        self._outline_locator_elapsed_ms = 0.0
        self._search_highlight: SearchMatch | None = None
        self._comment_marks: tuple[CommentMark, ...] = ()
        self._active_comment_id: str | None = None
        self._url_links = _page_url_links(page)
        self._text_selection: TextSelection | None = None
        self._selection_anchor: tuple[float, float] | None = None
        self._text_selection_start: tuple[float, float] | None = None
        self._text_selection_end: tuple[float, float] | None = None
        self._selection_handler = selection_handler
        self._selection_scroll_handler = selection_scroll_handler
        self._text_cursor: Gdk.Cursor | None = None
        self.set_app_paintable(True)
        self.set_halign(Gtk.Align.CENTER)
        _style(self, "see-docx-page")
        self._resize()
        self.connect("draw", self._on_draw)
        self.add_events(
            Gdk.EventMask.BUTTON_PRESS_MASK
            | Gdk.EventMask.BUTTON_RELEASE_MASK
            | Gdk.EventMask.BUTTON1_MOTION_MASK
            | Gdk.EventMask.POINTER_MOTION_MASK
            | Gdk.EventMask.SCROLL_MASK
            | Gdk.EventMask.ENTER_NOTIFY_MASK
            | Gdk.EventMask.LEAVE_NOTIFY_MASK
        )
        self.connect("button-press-event", self._on_button_press)
        self.connect("motion-notify-event", self._on_motion)
        self.connect("scroll-event", self._on_scroll)
        self.connect("button-release-event", self._on_button_release)
        self.connect("enter-notify-event", self._on_pointer_enter)
        self.connect("leave-notify-event", self._on_pointer_leave)

    def _resize(self) -> None:
        self.set_size_request(
            max(1, round(self._width * self._zoom)),
            max(1, round(self._height * self._zoom)),
        )

    def set_zoom(self, zoom: float) -> None:
        """Resize and redraw this existing page without replacing its widget."""

        if abs(zoom - self._zoom) < 0.001:
            return
        self._zoom = zoom
        self._resize()
        self.queue_draw()

    def destination_y(self, top: float) -> float:
        """Return the scroll Y coordinate for a PDF destination's top edge."""

        allocation = self.get_allocation()
        page_top = max(0.0, min(self._height, self._height - top))
        return allocation.y + page_top * self._zoom

    def url_link_position(self, link: UrlLink) -> tuple[int, int]:
        """Return a URI link's top-left corner in this widget's coordinates."""

        return (
            round(link.left * self._zoom),
            round((self._height - link.top) * self._zoom),
        )

    def url_link_size(self, link: UrlLink) -> tuple[int, int]:
        """Return the rendered width and height of a URI link's rectangle."""

        return (
            max(1, round((link.right - link.left) * self._zoom)),
            max(1, round((link.top - link.bottom) * self._zoom)),
        )

    def set_outline_locator(self, top: float, elapsed_ms: float) -> None:
        """Draw the current frame of a short-lived heading arrival marker."""

        if self._outline_locator_top != top:
            self._outline_locator_height = _outline_header_line_height(
                self._page, top
            )
        self._outline_locator_top = top
        self._outline_locator_elapsed_ms = max(elapsed_ms, 0.0)
        self.queue_draw()

    def clear_outline_locator(self) -> None:
        if self._outline_locator_top is None:
            return
        self._outline_locator_top = None
        self._outline_locator_height = OUTLINE_LOCATOR_FALLBACK_HEIGHT
        self._outline_locator_elapsed_ms = 0.0
        self.queue_draw()

    def set_search_highlight(self, match: SearchMatch) -> None:
        self._search_highlight = match
        self.queue_draw()

    def clear_search_highlight(self) -> None:
        if self._search_highlight is None:
            return
        self._search_highlight = None
        self.queue_draw()

    def set_comment_marks(
        self, marks: Iterable[CommentMark]
    ) -> None:
        """Paint the source text covered by comments without hiding glyphs."""

        self._comment_marks = tuple(marks)
        self.queue_draw()

    def set_active_comment(self, comment_id: str | None) -> None:
        """Set the comment range that should use the highlight role."""

        if self._active_comment_id == comment_id:
            return
        self._active_comment_id = comment_id
        self.queue_draw()

    def _page_point(
        self, event: Gdk.EventButton | Gdk.EventMotion
    ) -> tuple[float, float]:
        return (
            min(max(float(event.x) / self._zoom, 0.0), self._width),
            min(max(float(event.y) / self._zoom, 0.0), self._height),
        )

    @staticmethod
    def _raw_page_point(
        event: Gdk.EventButton | Gdk.EventMotion,
    ) -> tuple[float, float]:
        """Return a pointer location in unbounded rendered-page pixels."""

        return float(event.x), float(event.y)

    @staticmethod
    def _poppler_rectangle(selection: TextSelection) -> Poppler.Rectangle:
        rectangle = Poppler.Rectangle()
        rectangle.x1 = selection.left
        rectangle.y1 = selection.top
        rectangle.x2 = selection.right
        rectangle.y2 = selection.bottom
        return rectangle

    @staticmethod
    def _nearest_glyph_index(
        text: str,
        rectangles: list[Poppler.Rectangle],
        point: tuple[float, float],
    ) -> int | None:
        """Return the selectable glyph closest to a document-space pointer."""

        x, y = point

        def distance_to_interval(value: float, lower: float, upper: float) -> float:
            if value < lower:
                return lower - value
            if value > upper:
                return value - upper
            return 0.0

        candidates = (
            (
                distance_to_interval(y, float(rectangle.y1), float(rectangle.y2)),
                distance_to_interval(x, float(rectangle.x1), float(rectangle.x2)),
                index,
            )
            for index, (character, rectangle) in enumerate(
                zip(text, rectangles, strict=True)
            )
            if character not in "\r\n"
        )
        try:
            return min(candidates)[2]
        except ValueError:
            return None

    def _layout_text_selection(
        self, selection: TextSelection
    ) -> tuple[str, list[Poppler.Rectangle]] | None:
        """Return text and glyph bounds for the active rendered selection."""

        try:
            text = self._page.get_text()
            has_layout, rectangles = self._page.get_text_layout()
        except (AttributeError, TypeError, ValueError):
            return None
        if not has_layout or len(text) != len(rectangles):
            return None

        start = getattr(self, "_text_selection_start", None)
        end = getattr(self, "_text_selection_end", None)
        if start is not None and end is not None:
            first = self._nearest_glyph_index(text, rectangles, start)
            last = self._nearest_glyph_index(text, rectangles, end)
            if first is not None and last is not None:
                first, last = sorted((first, last))
                selected_rectangles = [
                    rectangle
                    for index, (character, rectangle) in enumerate(
                        zip(text, rectangles, strict=True)
                    )
                    if first <= index <= last and character not in "\r\n"
                ]
                return text[first : last + 1], selected_rectangles

        selected_text: list[str] = []
        selected_rectangles: list[Poppler.Rectangle] = []
        for character, rectangle in zip(text, rectangles, strict=True):
            if character in "\r\n":
                if selection.top <= float(rectangle.y1) <= selection.bottom:
                    selected_text.append(character)
                continue
            intersects = (
                float(rectangle.x2) >= selection.left
                and float(rectangle.x1) <= selection.right
                and float(rectangle.y2) >= selection.top
                and float(rectangle.y1) <= selection.bottom
            )
            if intersects:
                selected_text.append(character)
                selected_rectangles.append(rectangle)
        return "".join(selected_text), selected_rectangles

    def selected_text(self, selection: TextSelection) -> str:
        """Extract selected glyphs without Poppler's adjacent-line drift."""

        layout_selection = self._layout_text_selection(selection)
        if layout_selection is not None:
            text, _rectangles = layout_selection
        else:
            text = self._page.get_selected_text(
                Poppler.SelectionStyle.GLYPH,
                self._poppler_rectangle(selection),
            )
        return _reflow_pdf_selection_text(text)

    def selected_source_range(
        self, selection: TextSelection
    ) -> tuple[int, int] | None:
        """Return the OOXML source range represented by this glyph selection."""

        mapping = getattr(self, "_source_character_map", None)
        start = getattr(self, "_text_selection_start", None)
        end = getattr(self, "_text_selection_end", None)
        if not mapping or start is None or end is None:
            return None
        try:
            text = self._page.get_text()
            has_layout, rectangles = self._page.get_text_layout()
        except (AttributeError, TypeError, ValueError):
            return None
        if not has_layout or len(text) != len(rectangles):
            return None
        first = self._nearest_glyph_index(text, rectangles, start)
        last = self._nearest_glyph_index(text, rectangles, end)
        if first is None or last is None:
            return None
        source_indices = [
            mapping[index]
            for index in range(min(first, last), max(first, last) + 1)
            if index in mapping and text[index] not in "\r\n"
        ]
        if not source_indices:
            return None
        return min(source_indices), max(source_indices) + 1

    def _copy_selection(self, text: str, html: str | None = None) -> None:
        """Publish a selection for both Ctrl+V and primary paste.

        HTML is offered alongside text/plain so Writer receives formatted
        runs and table cells, while terminals and plain editors keep their
        expected UTF-8 fallback.
        """

        if not text:
            return
        for atom in (Gdk.SELECTION_CLIPBOARD, Gdk.SELECTION_PRIMARY):
            clipboard = Gtk.Clipboard.get(atom)
            if html and _publish_rich_clipboard(
                clipboard, RichTextClipboardPayload(text, html)
            ):
                continue
            clipboard.set_text(text, -1)

    def set_text_selection(
        self,
        selection: TextSelection | None,
        *,
        start: tuple[float, float] | None = None,
        end: tuple[float, float] | None = None,
    ) -> None:
        """Render one document-coordinated selection segment on this page."""

        self._text_selection = selection
        self._text_selection_start = start
        self._text_selection_end = end
        self.queue_draw()

    def clear_text_selection(self) -> None:
        """Remove this page's visible portion of a document selection."""

        if self._text_selection is None:
            return
        self._text_selection = None
        self._text_selection_start = None
        self._text_selection_end = None
        self.queue_draw()

    def _update_text_selection(self, point: tuple[float, float]) -> bool:
        if self._selection_anchor is None:
            return False
        self._text_selection_start = self._selection_anchor
        self._text_selection_end = point
        self._text_selection = _text_selection_bounds(
            start=self._selection_anchor,
            end=point,
            page_width=self._width,
            page_height=self._height,
            line_padding=TEXT_SELECTION_LINE_PADDING / self._zoom,
        )
        text = self.selected_text(self._text_selection)
        self._copy_selection(text)
        self.queue_draw()
        return True

    def _set_text_cursor(self) -> None:
        """Advertise that the rendered PDF glyphs can be selected."""

        page_window = self.get_window()
        if page_window is None:
            return
        if self._text_cursor is None:
            self._text_cursor = Gdk.Cursor.new_for_display(
                page_window.get_display(), Gdk.CursorType.XTERM
            )
        page_window.set_cursor(self._text_cursor)

    def _on_pointer_enter(
        self, _widget: Gtk.DrawingArea, _event: Gdk.EventCrossing
    ) -> bool:
        self._set_text_cursor()
        return False

    def _on_pointer_leave(
        self, _widget: Gtk.DrawingArea, _event: Gdk.EventCrossing
    ) -> bool:
        page_window = self.get_window()
        if page_window is not None:
            page_window.set_cursor(None)
        return False

    def _on_button_press(
        self, _widget: Gtk.DrawingArea, event: Gdk.EventButton
    ) -> bool:
        if event.button != 1:
            return False
        handler = getattr(self, "_selection_handler", None)
        if handler is not None:
            handler("begin", self, self._raw_page_point(event))
            return True
        self._selection_anchor = self._page_point(event)
        self._text_selection = None
        self.queue_draw()
        return True

    def _on_motion(self, _widget: Gtk.DrawingArea, event: Gdk.EventMotion) -> bool:
        self._set_text_cursor()
        handler = getattr(self, "_selection_handler", None)
        if handler is not None:
            handler("update", self, self._raw_page_point(event))
            return True
        return self._update_text_selection(self._page_point(event))

    def _on_scroll(self, _widget: Gtk.DrawingArea, event: Gdk.EventScroll) -> bool:
        """Route a held-drag wheel event to the document scroll container."""

        handler = getattr(self, "_selection_scroll_handler", None)
        return bool(handler(self, event)) if handler is not None else False

    def _on_button_release(
        self, _widget: Gtk.DrawingArea, event: Gdk.EventButton
    ) -> bool:
        handler = getattr(self, "_selection_handler", None)
        if handler is not None:
            if event.button != 1:
                return False
            handler("end", self, self._raw_page_point(event))
            return True
        if event.button != 1 or self._selection_anchor is None:
            return False
        self._update_text_selection(self._page_point(event))
        self._selection_anchor = None
        return True

    def _draw_outline_locator(self, context: object) -> None:
        if self._outline_locator_top is None:
            return
        frame = _outline_locator_frame(self._outline_locator_elapsed_ms)
        if frame.fill_opacity <= 0:
            return
        scale = self._zoom
        page_top = max(0.0, min(self._height, self._height - self._outline_locator_top))
        x = 0.0
        y = page_top * scale
        width = self._width * scale
        height = getattr(
            self, "_outline_locator_height", OUTLINE_LOCATOR_FALLBACK_HEIGHT
        ) * scale
        center_x = 72 * scale
        radius = height * 0.35 + width * 0.76 * frame.expansion
        palette = _theme_palette(self)
        accent = _rgb(palette["accent"])
        highlight = _rgb(palette["highlight"])

        # Keep the material-style ripple tied to the bookmark destination,
        # with a fixed origin at the heading margin. The square band clips
        # the expanding circle rather than moving it across the page.
        context.save()
        context.rectangle(x, y, width, height)
        context.clip()
        # Multiply tints the rendered page fill without lightening or
        # obscuring its black PDF glyphs, so the locator reads as a true
        # background highlight rather than a layer over the heading.
        context.set_operator(cairo.OPERATOR_MULTIPLY)
        context.set_source_rgba(*accent, 0.025 * frame.fill_opacity)
        context.rectangle(x, y, width, height)
        context.fill()
        context.set_source_rgba(*accent, 0.22 * frame.fill_opacity)
        context.arc(center_x, y + height / 2, radius, 0, 2 * math.pi)
        context.fill()
        context.set_source_rgba(*highlight, 0.14 * frame.bloom_opacity)
        context.arc(center_x, y + height / 2, radius * 0.46, 0, 2 * math.pi)
        context.fill()
        context.restore()

    def _draw_search_highlight(self, context: object) -> None:
        match = self._search_highlight
        if match is None:
            return
        x = match.left * self._zoom
        y = (self._height - match.top) * self._zoom
        width = max(1.0, (match.right - match.left) * self._zoom)
        height = max(1.0, (match.top - match.bottom) * self._zoom)
        palette = _theme_palette(self)
        accent = _rgb(palette["accent"])
        highlight = _rgb(palette["highlight"])
        context.set_source_rgba(*accent, 0.28)
        context.rectangle(x - 2, y - 2, width + 4, height + 4)
        context.fill_preserve()
        context.set_source_rgba(*highlight, 0.96)
        context.set_line_width(max(1.0, 1.25 * self._zoom))
        context.stroke()

    def _draw_comment_marks(self, context: object) -> None:
        """Keep context quiet and make the focused range legible on paper.

        SC1's highlight role is intentionally light for dark UI surfaces. On
        the white PDF sheet it has almost no contrast, so the page-specific
        focus treatment uses the saturated variant accent instead.
        """

        if not self._comment_marks:
            return
        palette = _theme_palette(self)
        context.save()
        inactive_rectangles = _coalesce_comment_rectangles(
            mark.rectangle
            for mark in self._comment_marks
            if mark.comment_id != self._active_comment_id
        )
        active_rectangles = _coalesce_comment_rectangles(
            mark.rectangle
            for mark in self._comment_marks
            if mark.comment_id == self._active_comment_id
        )
        for rectangles, active in (
            (inactive_rectangles, False),
            (active_rectangles, True),
        ):
            color = _rgb(palette["accent"])
            fill_alpha = 0.03 if not active else 0.34
            line_alpha = 0.18 if not active else 0.99
            line_width = max(1.0, self._zoom) if not active else max(
                2.25, 2.25 * self._zoom
            )
            for left, top, right, bottom in rectangles:
                x = left * self._zoom
                y = top * self._zoom
                width = max(1.0, (right - left) * self._zoom)
                height = max(1.0, (bottom - top) * self._zoom)
                context.set_operator(
                    cairo.OPERATOR_OVER if active else cairo.OPERATOR_MULTIPLY
                )
                context.set_source_rgba(*color, fill_alpha)
                context.rectangle(x, y, width, height)
                context.fill()
                context.set_operator(cairo.OPERATOR_OVER)
                context.set_source_rgba(*color, line_alpha)
                context.set_line_width(line_width)
                context.move_to(x, y + height + 1.0)
                context.line_to(x + width, y + height + 1.0)
                context.stroke()
        context.restore()

    def _draw_text_selection(self, context: object) -> None:
        selection = self._text_selection
        if selection is None:
            return
        layout_selection = self._layout_text_selection(selection)
        context.set_source_rgba(0.25, 0.58, 0.92, 0.32)
        if layout_selection is not None:
            _text, rectangles = layout_selection
            for rectangle in rectangles:
                context.rectangle(
                    float(rectangle.x1) * self._zoom,
                    float(rectangle.y1) * self._zoom,
                    (float(rectangle.x2) - float(rectangle.x1)) * self._zoom,
                    (float(rectangle.y2) - float(rectangle.y1)) * self._zoom,
                )
        else:
            region = self._page.get_selected_region(
                self._zoom,
                Poppler.SelectionStyle.GLYPH,
                self._poppler_rectangle(selection),
            )
            for index in range(region.num_rectangles()):
                rectangle = region.get_rectangle(index)
                context.rectangle(
                    rectangle.x,
                    rectangle.y,
                    rectangle.width,
                    rectangle.height,
                )
        context.fill()

    def _on_draw(self, _widget: Gtk.DrawingArea, context: object) -> bool:
        # Poppler receives the Pycairo context supplied by GTK.
        # DrawingArea CSS is not guaranteed to supply an opaque surface, so
        # paint the document sheet explicitly before the PDF's black text.
        context.set_source_rgb(1.0, 1.0, 1.0)
        context.paint()
        context.save()
        context.scale(self._zoom, self._zoom)
        self._page.render(context)
        context.restore()
        self._draw_outline_locator(context)
        self._draw_search_highlight(context)
        self._draw_comment_marks(context)
        self._draw_text_selection(context)
        return False


class PdfDocumentView:
    """A scrollable PDF page stack that can restore a document location."""

    def __init__(
        self,
        on_page_changed: Callable[[int | None, int], None] | None = None,
    ) -> None:
        self.widget = Gtk.ScrolledWindow()
        # Keep GTK's vertical adjustment active for wheel, touchpad, keyboard,
        # and programmatic navigation.  The corresponding scrollbar chrome is
        # hidden in CSS, rather than disabling the scrolling policy itself.
        self.widget.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        self.widget.set_kinetic_scrolling(True)
        self.widget.add_events(Gdk.EventMask.SCROLL_MASK)
        self.widget.connect("scroll-event", self._on_scroll_event)
        _style(self.widget, "see-docx-workspace")

        # The gap is part of the GTK layout, rather than CSS decoration:
        # GTK3 does not consistently allocate CSS margins between children.
        # That makes each PDF sheet visibly distinct in every theme.
        self._pages_box = Gtk.Box(
            orientation=Gtk.Orientation.VERTICAL, spacing=PAGE_GAP
        )
        self._pages_box.set_halign(Gtk.Align.CENTER)
        self._pages_box.set_valign(Gtk.Align.START)
        self._pages_box.set_margin_top(PAGE_MARGIN)
        self._pages_box.set_margin_bottom(PAGE_MARGIN)
        self._pages_box.set_margin_start(PAGE_MARGIN)
        self._pages_box.set_margin_end(PAGE_MARGIN)
        _style(self._pages_box, "see-docx-pages")
        self.widget.add(self._pages_box)
        self._pages: list[PdfPage] = []
        self._document: Poppler.Document | None = None
        self._rich_source: _DocxRichTextSource | None = None
        self._source_text: str | None = None
        self._comment_anchors: tuple[CommentAnchor, ...] = ()
        self._active_comment_id: str | None = None
        self._outline: list[OutlineEntry] = []
        self._outline_locator_source = 0
        self._outline_locator_page: PdfPage | None = None
        self._search_highlight_page: PdfPage | None = None
        self._selection_anchor: tuple[float, float] | None = None
        self._selection_endpoint: tuple[float, float] | None = None
        self._selection_anchor_page: PdfPage | None = None
        self._selection_drag_active = False
        self._selection_auto_scroll_source = 0
        self._selection_auto_scroll_page: PdfPage | None = None
        self._selection_auto_scroll_point: tuple[float, float] | None = None
        self._pending_restore: DocumentPosition | None = None
        self._restore_attempts = 0
        self._on_page_changed = on_page_changed
        self.zoom = DEFAULT_ZOOM
        adjustment = self.widget.get_vadjustment()
        self._selection_scroll_value = adjustment.get_value()
        adjustment.connect("changed", self._on_scroll_changed)
        adjustment.connect("value-changed", self._on_scroll_changed)

    @property
    def has_document(self) -> bool:
        return self._document is not None

    @property
    def page_count(self) -> int:
        return len(self._pages)

    @property
    def current_page_index(self) -> int | None:
        """The zero-based page at the viewport's current reading location."""

        adjustment = self.widget.get_vadjustment()
        return page_index_at_scroll(
            adjustment.get_value() - adjustment.get_lower(),
            self._page_geometries(),
            maximum_scroll=self._maximum_scroll(),
        )

    @property
    def outline(self) -> tuple[OutlineEntry, ...]:
        """The document headings available for sidebar navigation."""

        return tuple(self._outline)

    @property
    def comments(self) -> tuple[DocumentComment, ...]:
        """The source comments that belong to the current rendered document."""

        return self._rich_source.comments if self._rich_source is not None else ()

    @property
    def comment_threads(self) -> tuple[CommentThread, ...]:
        """The anchored conversations that belong to the current document."""

        return _comment_threads(self.comments)

    @property
    def comment_anchors(self) -> tuple[CommentAnchor, ...]:
        """The rendered locations available for the comment connector layer."""

        return self._comment_anchors

    def set_active_comment(self, comment_id: str | None) -> None:
        """Color only the focused comment's source range as highlighted."""

        self._active_comment_id = comment_id
        for page in self._pages:
            page.set_active_comment(comment_id)

    def comment_anchor_text(self, comment: DocumentComment) -> str:
        """Return the source text a comment is attached to for the rail quote."""

        source = self._rich_source
        if source is None:
            return ""
        return source._selection_text[comment.source_start : comment.source_end]

    def _on_scroll_changed(self, _adjustment: Gtk.Adjustment) -> None:
        current_scroll = _adjustment.get_value()
        previous_scroll = getattr(self, "_selection_scroll_value", current_scroll)
        self._selection_scroll_value = current_scroll
        if (
            self._selection_drag_active
            and self._selection_endpoint is not None
            and current_scroll != previous_scroll
        ):
            endpoint_x, endpoint_y = self._selection_endpoint
            self._selection_endpoint = (
                endpoint_x,
                endpoint_y + current_scroll - previous_scroll,
            )
            self._apply_document_selection()
        self._notify_page_changed()

    def _on_scroll_event(
        self, _widget: Gtk.ScrolledWindow, event: Gdk.EventScroll
    ) -> bool:
        """Zoom with Ctrl+wheel while preserving ordinary document scrolling."""

        if not event.state & Gdk.ModifierType.CONTROL_MASK:
            return False
        if event.direction == Gdk.ScrollDirection.UP:
            increment = ZOOM_STEP
        elif event.direction == Gdk.ScrollDirection.DOWN:
            increment = -ZOOM_STEP
        elif event.direction == Gdk.ScrollDirection.SMOOTH:
            has_deltas, _delta_x, delta_y = event.get_scroll_deltas()
            if not has_deltas or delta_y == 0:
                return False
            increment = -ZOOM_STEP if delta_y > 0 else ZOOM_STEP
        else:
            return False
        self.set_zoom(self.zoom + increment)
        return True

    def _document_point_for_page(
        self, page: PdfPage, point: tuple[float, float]
    ) -> tuple[float, float]:
        """Convert rendered-page coordinates to unscrolled document space."""

        allocation = page.get_allocation()
        return float(allocation.x) + point[0], float(allocation.y) + point[1]

    def _page_point_for_document(
        self, page: PdfPage, point: tuple[float, float]
    ) -> tuple[float, float]:
        """Convert unscrolled document coordinates to one PDF page's space."""

        allocation = page.get_allocation()
        return (
            (point[0] - float(allocation.x)) / page._zoom,
            (point[1] - float(allocation.y)) / page._zoom,
        )

    def _selection_page_index(self, point: tuple[float, float]) -> int | None:
        """Return the nearest page at a document-space pointer position."""

        if not self._pages:
            return None
        point_y = point[1]
        for index, page in enumerate(self._pages):
            allocation = page.get_allocation()
            if point_y <= float(allocation.y) + float(allocation.height):
                return index
        return len(self._pages) - 1

    @staticmethod
    def _selection_is_forward(
        anchor: tuple[float, float], endpoint: tuple[float, float]
    ) -> bool:
        return endpoint[1] > anchor[1] or (
            endpoint[1] == anchor[1] and endpoint[0] >= anchor[0]
        )

    def _clear_page_text_selections(self) -> None:
        for page in self._pages:
            page.clear_text_selection()

    def _clear_document_selection(self) -> None:
        """End any held drag and remove its persistent visual selection."""

        self._stop_selection_auto_scroll()
        self._selection_anchor = None
        self._selection_endpoint = None
        self._selection_anchor_page = None
        self._selection_drag_active = False
        self._clear_page_text_selections()

    def cancel_text_selection(self) -> None:
        """Clear pointer selection before document replacement or shutdown."""

        self._clear_document_selection()

    def _on_page_selection_event(
        self, phase: str, page: PdfPage, point: tuple[float, float]
    ) -> None:
        """Coordinate one pointer drag across the complete PDF document."""

        document_point = self._document_point_for_page(page, point)
        if phase == "begin":
            self._clear_document_selection()
            self._selection_anchor = document_point
            self._selection_anchor_page = page
            self._selection_drag_active = True
        elif not self._selection_drag_active:
            return
        self._selection_endpoint = document_point
        self._apply_document_selection()
        if phase == "update":
            self._update_selection_auto_scroll(page, point)
        elif phase == "end":
            self._selection_drag_active = False
            self._stop_selection_auto_scroll()

    def _on_page_selection_scroll_event(
        self, _page: PdfPage, event: Gdk.EventScroll
    ) -> bool:
        """Scroll a held selection whose wheel event targets a PDF page."""

        if event.state & Gdk.ModifierType.CONTROL_MASK:
            return self._on_scroll_event(self.widget, event)
        if not self._selection_drag_active:
            return False
        if event.direction == Gdk.ScrollDirection.UP:
            self.scroll("line-up")
            return True
        if event.direction == Gdk.ScrollDirection.DOWN:
            self.scroll("line-down")
            return True
        if event.direction != Gdk.ScrollDirection.SMOOTH:
            return False
        _success, _delta_x, delta_y = event.get_scroll_deltas()
        if not delta_y:
            return False
        adjustment = self.widget.get_vadjustment()
        lower = adjustment.get_lower()
        target = adjustment.get_value() + delta_y * SCROLL_STEP
        adjustment.set_value(min(max(target, lower), lower + self._maximum_scroll()))
        return True

    def _selection_regions(
        self,
    ) -> list[
        tuple[
            PdfPage,
            TextSelection,
            tuple[float, float],
            tuple[float, float],
        ]
    ]:
        """Split the document-level drag into Poppler regions per page."""

        anchor = self._selection_anchor
        endpoint = self._selection_endpoint
        if anchor is None or endpoint is None or anchor == endpoint:
            return []
        anchor_index = self._selection_page_index(anchor)
        endpoint_index = self._selection_page_index(endpoint)
        if anchor_index is None or endpoint_index is None:
            return []
        if self._selection_is_forward(anchor, endpoint):
            first_index, first_point = anchor_index, anchor
            last_index, last_point = endpoint_index, endpoint
        else:
            first_index, first_point = endpoint_index, endpoint
            last_index, last_point = anchor_index, anchor

        regions: list[
            tuple[
                PdfPage,
                TextSelection,
                tuple[float, float],
                tuple[float, float],
            ]
        ] = []
        for index in range(first_index, last_index + 1):
            page = self._pages[index]
            if first_index == last_index:
                start = self._page_point_for_document(page, first_point)
                end = self._page_point_for_document(page, last_point)
            elif index == first_index:
                start = self._page_point_for_document(page, first_point)
                end = (page._width, page._height)
            elif index == last_index:
                start = (0.0, 0.0)
                end = self._page_point_for_document(page, last_point)
            else:
                start = (0.0, 0.0)
                end = (page._width, page._height)
            regions.append(
                (
                    page,
                    _text_selection_bounds(
                        start=start,
                        end=end,
                        page_width=page._width,
                        page_height=page._height,
                        line_padding=TEXT_SELECTION_LINE_PADDING / page._zoom,
                    ),
                    start,
                    end,
                )
            )
        return regions

    def _map_pdf_glyphs_to_source(self) -> None:
        """Map rendered PDF glyph positions back to their OOXML characters."""

        source = self._rich_source
        if source is None:
            return
        page_text: list[str] = []
        offsets: list[int] = []
        cursor = 0
        for index, page in enumerate(self._pages):
            try:
                text = page._page.get_text()
            except (AttributeError, TypeError, ValueError):
                text = ""
            page._source_character_map = {}
            offsets.append(cursor)
            page_text.append(text)
            cursor += len(text)
            if index + 1 < len(self._pages):
                # A page boundary is always at least a visual text boundary.
                # Folding it to whitespace keeps wrapped paragraphs aligned.
                cursor += 1
        rendered_text = "\n".join(page_text)
        rendered, rendered_indices = _fold_selection_whitespace(rendered_text)
        source_text, source_indices = _fold_selection_whitespace(source._selection_text)
        if not rendered or not source_text:
            return
        matcher = SequenceMatcher(None, rendered, source_text)
        for match in matcher.get_matching_blocks():
            for delta in range(match.size):
                rendered_index = rendered_indices[match.a + delta]
                page_index = 0
                while (
                    page_index + 1 < len(offsets)
                    and rendered_index >= offsets[page_index + 1]
                ):
                    page_index += 1
                local_index = rendered_index - offsets[page_index]
                if 0 <= local_index < len(page_text[page_index]):
                    self._pages[page_index]._source_character_map[local_index] = (
                        source_indices[match.b + delta]
                    )

    def _map_comments_to_pages(self) -> None:
        """Map OOXML comment ranges onto the PDF glyph rectangles they cover."""

        marks_by_page: dict[int, list[CommentMark]] = {
            index: [] for index in range(len(self._pages))
        }
        anchors: list[CommentAnchor] = []
        for comment in self.comments:
            matches_by_page: dict[
                int, list[tuple[float, float, float, float]]
            ] = {}
            for page_index, page in enumerate(self._pages):
                mapping = getattr(page, "_source_character_map", {})
                if not mapping:
                    continue
                try:
                    text = page._page.get_text()
                    has_layout, rectangles = page._page.get_text_layout()
                except (AttributeError, TypeError, ValueError):
                    continue
                if not has_layout or len(text) != len(rectangles):
                    continue
                page_rectangles = []
                for rendered_index, source_index in mapping.items():
                    if not comment.source_start <= source_index < comment.source_end:
                        continue
                    if not 0 <= rendered_index < len(text) or text[rendered_index].isspace():
                        continue
                    rectangle = rectangles[rendered_index]
                    page_rectangles.append(
                        (
                            min(float(rectangle.x1), float(rectangle.x2)),
                            min(float(rectangle.y1), float(rectangle.y2)),
                            max(float(rectangle.x1), float(rectangle.x2)),
                            max(float(rectangle.y1), float(rectangle.y2)),
                        )
                    )
                if page_rectangles:
                    matches_by_page[page_index] = page_rectangles
                    marks_by_page[page_index].extend(
                        CommentMark(comment.comment_id, rectangle)
                        for rectangle in page_rectangles
                    )
            if matches_by_page:
                page_index = min(matches_by_page)
                anchors.append(
                    CommentAnchor(
                        comment=comment,
                        page_index=page_index,
                        rectangles=tuple(matches_by_page[page_index]),
                    )
                )
        self._comment_anchors = tuple(anchors)
        for page_index, page in enumerate(self._pages):
            page.set_comment_marks(marks_by_page[page_index])

    def _apply_document_selection(self) -> None:
        """Render and copy every page segment covered by the active drag."""

        regions = self._selection_regions()
        selected_page_ids = {
            id(page) for page, _selection, _start, _end in regions
        }
        for page in self._pages:
            if id(page) not in selected_page_ids:
                page.clear_text_selection()

        text: list[str] = []
        source_ranges: list[tuple[int, int]] = []
        for page, selection, start, end in regions:
            page.set_text_selection(selection, start=start, end=end)
            selected = page.selected_text(selection)
            if selected:
                text.append(selected)
            source_range = page.selected_source_range(selection)
            if source_range is not None:
                source_ranges.append(source_range)
        if text and self._selection_anchor_page is not None:
            selected_text = "\n".join(text)
            source = getattr(self, "_rich_source", None)
            payload = None
            if source is not None and source_ranges:
                payload = source.payload_for_range(
                    min(start for start, _end in source_ranges),
                    max(end for _start, end in source_ranges),
                )
            if payload is None and source is not None:
                payload = source.payload_for_text(selected_text)
            if payload is not None:
                self._selection_anchor_page._copy_selection(payload.text, payload.html)
                return
            source_text = getattr(self, "_source_text", None)
            if source_text is not None:
                selected_text = (
                    _restore_docx_paragraph_boundaries(source_text, selected_text)
                    or selected_text
                )
            self._selection_anchor_page._copy_selection(selected_text)

    def _update_selection_auto_scroll(
        self, page: PdfPage, point: tuple[float, float]
    ) -> None:
        """Continue a drag when its pointer moves above or below a page."""

        page_height = float(page.get_allocated_height())
        if 0.0 <= point[1] <= page_height:
            self._stop_selection_auto_scroll()
            return
        self._selection_auto_scroll_page = page
        self._selection_auto_scroll_point = point
        if not self._selection_auto_scroll_source:
            self._selection_auto_scroll_source = GLib.timeout_add(
                SELECTION_AUTO_SCROLL_TICK_MS, self._auto_scroll_selection
            )

    def _stop_selection_auto_scroll(self) -> None:
        if self._selection_auto_scroll_source:
            GLib.source_remove(self._selection_auto_scroll_source)
            self._selection_auto_scroll_source = 0
        self._selection_auto_scroll_page = None
        self._selection_auto_scroll_point = None

    def _auto_scroll_selection(self) -> bool:
        """Advance a held edge drag at a speed proportional to its overflow."""

        page = self._selection_auto_scroll_page
        point = self._selection_auto_scroll_point
        if not self._selection_drag_active or page is None or point is None:
            self._selection_auto_scroll_source = 0
            return GLib.SOURCE_REMOVE
        page_height = float(page.get_allocated_height())
        if point[1] < 0.0:
            direction, overflow = -1.0, -point[1]
        elif point[1] > page_height:
            direction, overflow = 1.0, point[1] - page_height
        else:
            self._stop_selection_auto_scroll()
            return GLib.SOURCE_REMOVE
        step = min(
            SELECTION_AUTO_SCROLL_MAX_STEP,
            SELECTION_AUTO_SCROLL_MIN_STEP + overflow * 0.25,
        )
        adjustment = self.widget.get_vadjustment()
        value = adjustment.get_value()
        lower = adjustment.get_lower()
        upper = lower + self._maximum_scroll()
        target = min(max(value + direction * step, lower), upper)
        if target == value:
            self._stop_selection_auto_scroll()
            return GLib.SOURCE_REMOVE
        adjustment.set_value(target)
        return GLib.SOURCE_CONTINUE

    def _notify_page_changed(self) -> None:
        if self._on_page_changed is not None:
            self._on_page_changed(self.current_page_index, self.page_count)

    def set_document(
        self,
        document: Poppler.Document,
        *,
        source: _DocxRichTextSource | None = None,
        source_text: str | None = None,
    ) -> None:
        self._clear_outline_locator()
        self.clear_search_highlight()
        self._clear_document_selection()
        for page in self._pages:
            self._pages_box.remove(page)
        self._pages.clear()
        self._document = document
        self._active_comment_id = None
        self._rich_source = source
        self._source_text = source.plain_text if source is not None else source_text
        self._outline = _document_outline(document)
        for number in range(document.get_n_pages()):
            page = PdfPage(
                document.get_page(number),
                self.zoom,
                selection_handler=self._on_page_selection_event,
                selection_scroll_handler=self._on_page_selection_scroll_event,
            )
            self._pages_box.pack_start(page, False, False, 0)
            self._pages.append(page)
        self._map_pdf_glyphs_to_source()
        self._map_comments_to_pages()
        self._pages_box.show_all()
        self._notify_page_changed()

    def set_zoom(self, zoom: float, *, minimum: float = MIN_ZOOM) -> bool:
        if self._document is None:
            return False
        zoom = min(max(zoom, minimum), MAX_ZOOM)
        if abs(zoom - self.zoom) < 0.001:
            return False
        position = self.capture_position()
        self.zoom = zoom
        for page in self._pages:
            page.set_zoom(zoom)
        self._pages_box.queue_resize()
        GLib.idle_add(self.restore_position_after_layout, position)
        return True

    def fit_to_viewport(self) -> bool:
        """Zoom out only as far as required to keep a complete page visible."""

        if not self._pages:
            return False
        target = _fit_zoom_for_viewport(
            self._pages[0]._width,
            self._pages[0]._height,
            self.widget.get_allocated_width(),
            self.widget.get_allocated_height(),
        )
        return self.set_zoom(min(self.zoom, target), minimum=MIN_FIT_ZOOM)

    def fit_to_width(self) -> bool:
        """Zoom out enough to keep the page width visible while scrolling vertically."""

        if not self._pages:
            return False
        viewport_width = self.widget.get_allocated_width()
        target = max(
            MIN_FIT_ZOOM,
            (viewport_width - 2 * PAGE_MARGIN) / self._pages[0]._width,
        )
        return self.set_zoom(min(self.zoom, target), minimum=MIN_FIT_ZOOM)

    def _page_geometries(self) -> list[PageGeometry]:
        geometries: list[PageGeometry] = []
        for page in self._pages:
            allocation = page.get_allocation()
            geometries.append(
                PageGeometry(float(allocation.y), float(allocation.height))
            )
        return geometries

    def _maximum_scroll(self) -> float:
        adjustment = self.widget.get_vadjustment()
        return max(
            0.0,
            adjustment.get_upper()
            - adjustment.get_lower()
            - adjustment.get_page_size(),
        )

    def capture_position(self) -> DocumentPosition:
        adjustment = self.widget.get_vadjustment()
        return capture_position(
            adjustment.get_value() - adjustment.get_lower(),
            self._page_geometries(),
            maximum_scroll=self._maximum_scroll(),
        )

    def restore_position(self, position: DocumentPosition) -> bool:
        adjustment = self.widget.get_vadjustment()
        target = restore_position(
            position,
            self._page_geometries(),
            maximum_scroll=self._maximum_scroll(),
        )
        adjustment.set_value(target + adjustment.get_lower())
        self._notify_page_changed()
        return GLib.SOURCE_REMOVE

    def restore_position_after_layout(self, position: DocumentPosition) -> bool:
        """Restore only after GTK has assigned a real scroll range to pages."""

        self._pending_restore = position
        self._restore_attempts = 0
        GLib.timeout_add(25, self._restore_when_ready)
        return GLib.SOURCE_REMOVE

    @property
    def restore_pending(self) -> bool:
        return self._pending_restore is not None

    def _restore_when_ready(self) -> bool:
        position = self._pending_restore
        if position is None:
            return GLib.SOURCE_REMOVE
        self._restore_attempts += 1
        has_allocated_pages = bool(self._pages) and all(
            page.get_allocated_height() > 1 for page in self._pages
        )
        if self._maximum_scroll() <= 0 or not has_allocated_pages:
            if self._restore_attempts < 40:
                return GLib.SOURCE_CONTINUE
            self._pending_restore = None
            return GLib.SOURCE_REMOVE
        self._pending_restore = None
        return self.restore_position(position)

    def scroll(self, command: str) -> None:
        adjustment = self.widget.get_vadjustment()
        lower = adjustment.get_lower()
        maximum = lower + self._maximum_scroll()
        value = adjustment.get_value()
        if command == "line-down":
            value += SCROLL_STEP
        elif command == "line-up":
            value -= SCROLL_STEP
        elif command == "half-down":
            value += adjustment.get_page_size() / 2
        elif command == "half-up":
            value -= adjustment.get_page_size() / 2
        elif command == "top":
            value = lower
        elif command == "bottom":
            value = maximum
        adjustment.set_value(min(max(value, lower), maximum))

    def go_to_page(self, page_index: int) -> bool:
        """Place the selected page at the top of the viewport when possible."""

        if not 0 <= page_index < self.page_count:
            return False
        adjustment = self.widget.get_vadjustment()
        lower = adjustment.get_lower()
        target = self._pages[page_index].get_allocation().y
        target = min(max(target, 0.0), self._maximum_scroll())
        adjustment.set_value(target + lower)
        self._notify_page_changed()
        return True

    def go_to_adjacent_page(self, direction: int) -> bool:
        """Navigate one document page backward or forward without wrapping."""

        page_index = self.current_page_index
        if page_index is None:
            return False
        return self.go_to_page(page_index + direction)

    def go_to_outline_entry(self, entry: OutlineEntry) -> bool:
        """Navigate to an exported heading destination without changing zoom."""

        if not 0 <= entry.page_index < self.page_count:
            return False
        adjustment = self.widget.get_vadjustment()
        target = self._pages[entry.page_index].get_allocation().y
        if entry.top is not None:
            destination = self._pages[entry.page_index].destination_y(entry.top)
            target = _contextual_scroll_target(
                destination, adjustment.get_page_size(), self._maximum_scroll()
            )
            self._show_outline_locator(self._pages[entry.page_index], entry.top)
        else:
            self._clear_outline_locator()
        target = min(max(target, 0.0), self._maximum_scroll())
        adjustment.set_value(target + adjustment.get_lower())
        self._notify_page_changed()
        return True

    def go_to_comment_anchor(self, anchor: CommentAnchor) -> bool:
        """Bring a comment's attached text into the reading context."""

        if not 0 <= anchor.page_index < self.page_count:
            return False
        page = self._pages[anchor.page_index]
        adjustment = self.widget.get_vadjustment()
        destination = page.get_allocation().y + anchor.center_y * page._zoom
        target = destination - adjustment.get_page_size() / 2
        target = min(max(target, 0.0), self._maximum_scroll())
        adjustment.set_value(target + adjustment.get_lower())
        self._notify_page_changed()
        return True

    def _show_outline_locator(self, page: PdfPage, top: float) -> None:
        self._clear_outline_locator()
        self._outline_locator_page = page
        started = GLib.get_monotonic_time()
        page.set_outline_locator(top, 0.0)

        def fade() -> bool:
            elapsed_ms = (GLib.get_monotonic_time() - started) / 1_000
            if elapsed_ms >= OUTLINE_LOCATOR_DURATION_MS:
                page.clear_outline_locator()
                self._outline_locator_page = None
                self._outline_locator_source = 0
                return GLib.SOURCE_REMOVE
            page.set_outline_locator(top, elapsed_ms)
            return GLib.SOURCE_CONTINUE

        self._outline_locator_source = GLib.timeout_add(OUTLINE_LOCATOR_TICK_MS, fade)

    def _clear_outline_locator(self) -> None:
        if self._outline_locator_source:
            GLib.source_remove(self._outline_locator_source)
            self._outline_locator_source = 0
        if self._outline_locator_page is not None:
            self._outline_locator_page.clear_outline_locator()
            self._outline_locator_page = None

    def search(self, query: str) -> list[SearchMatch]:
        """Return native PDF text matches for the document, in reading order."""

        if self._document is None:
            return []
        return _document_search(self._document, query)

    def go_to_search_match(self, match: SearchMatch) -> bool:
        """Place a text match near the top of the viewport."""

        target = self._search_match_scroll_target(match)
        if target is None:
            return False
        adjustment = self.widget.get_vadjustment()
        self._set_search_highlight(self._pages[match.page_index], match)
        adjustment.set_value(target + adjustment.get_lower())
        self._notify_page_changed()
        return True

    def search_match_progress_fraction(self, match: SearchMatch) -> float | None:
        """Return the progress position reached when *match* is selected."""

        target = self._search_match_scroll_target(match)
        if target is None:
            return None
        return _reading_progress_fraction(target, self._maximum_scroll())

    def _search_match_scroll_target(self, match: SearchMatch) -> float | None:
        """Return the clamped viewport position used for a search match."""

        if not 0 <= match.page_index < self.page_count:
            return None
        adjustment = self.widget.get_vadjustment()
        destination = self._pages[match.page_index].destination_y(match.top)
        return _search_scroll_target(
            destination, adjustment.get_page_size(), self._maximum_scroll()
        )

    def _set_search_highlight(self, page: PdfPage, match: SearchMatch) -> None:
        self.clear_search_highlight()
        self._search_highlight_page = page
        page.set_search_highlight(match)

    def clear_search_highlight(self) -> None:
        if self._search_highlight_page is not None:
            self._search_highlight_page.clear_search_highlight()
            self._search_highlight_page = None

    def copy_all_text(self) -> bool:
        """Copy the document without turning every page into a highlight."""

        source = getattr(self, "_rich_source", None)
        if source is not None and self._pages:
            payload = source.full_payload()
            if payload.text:
                self._pages[0]._copy_selection(payload.text, payload.html)
                return True
        text = self._source_text
        if text is None:
            page_text: list[str] = []
            for page in self._pages:
                try:
                    extracted = page._page.get_text()
                except (AttributeError, TypeError, ValueError):
                    continue
                reflowed = _reflow_pdf_selection_text(extracted)
                if reflowed:
                    page_text.append(reflowed)
            text = "\n".join(page_text)
        if not text or not self._pages:
            return False
        self._pages[0]._copy_selection(text)
        return True


class DocxWindow(Gtk.ApplicationWindow):
    """A read-only, self-refreshing DOCX preview window."""

    def __init__(self, application: Gtk.Application, path: Path) -> None:
        super().__init__(application=application)
        self.path = path.expanduser().resolve()
        self._converter = LibreOfficeConverter()
        self._monitor: Gio.FileMonitor | None = None
        self._process: Gio.Subprocess | None = None
        self._export_converter: LibreOfficeConverter | PandocConverter | None = None
        self._export_process: Gio.Subprocess | None = None
        self._export_paths: ConversionPaths | PandocConversionPaths | None = None
        self._export_destination: Path | None = None
        self._writer_handoff_process: Gio.Subprocess | None = None
        self._active_paths: ConversionPaths | None = None
        self._active_revision: int | None = None
        self._revision = 0
        self._rendered_revision = 0
        self._debounce_source = 0
        self._closed = False
        self._last_status = "Preparing live preview…"
        self._pending_g = False
        self._url_hint_targets: dict[str, str] = {}
        self._url_hint_labels: dict[str, Gtk.Label] = {}
        self._url_hint_prefix = ""
        self._outline_count = 0
        self._page_indicator: Gtk.Label
        self._outline_store: Gtk.TreeStore
        self._outline_tree: Gtk.TreeView
        self._outline_panel: Gtk.Box
        self._outline_empty: Gtk.Label
        self._outline_entries: list[OutlineEntry] = []
        self._outline_row_paths: list[Gtk.TreePath] = []
        self._outline_zoom_before_open: float | None = None
        self._outline_fit_source = 0
        self._search_panel: Gtk.Box
        self._search_entry: Gtk.Entry
        self._search_status: Gtk.Label
        self._search_matches: list[SearchMatch] = []
        self._search_index = -1
        self._search_session_committed = False
        self._search_session_status: Gtk.Label
        self._search_match_marker_layer: Gtk.DrawingArea
        self._page_jump_panel: Gtk.Box
        self._page_jump_entry: Gtk.Entry
        self._page_jump_status: Gtk.Label
        self._export_panel: Gtk.Box
        self._export_list: Gtk.ListBox
        self._export_status: Gtk.Label
        self._export_index = 0
        self._comments_revealer: Gtk.Revealer
        self._comments_panel: Gtk.Box
        self._comments_scroller: Gtk.ScrolledWindow
        self._comments_list: Gtk.Box
        self._comments_empty: Gtk.Label
        self._comments_summary: Gtk.Label
        self._comments_count: Gtk.Label
        self._comment_line_layer: Gtk.DrawingArea
        self._comment_cards: dict[str, Gtk.Box] = {}
        self._comment_body_scrollers: dict[str, Gtk.ScrolledWindow] = {}
        self._comment_annotations: tuple[DocumentComment, ...] = ()
        self._comment_anchor_by_id: dict[str, CommentAnchor] = {}
        self._active_comment_index = -1
        self._active_comment_id: str | None = None
        self._comments_available = False
        self._comments_visible = False
        self._comments_focused = False
        self._comment_body_focused = False
        self._url_hint_layer: Gtk.Fixed
        self._path_status: Gtk.Label
        self._reading_progress_source = 0
        self._reading_progress_start = 0.0
        self._reading_progress_target = 0.0
        self._reading_progress_started_at = 0

        self.set_default_size(1060, 760)
        self.set_title(f"{self.path.name} — See DOCX")
        self.set_icon_name("x-office-document")
        self.connect("delete-event", self._on_delete)
        self.connect("key-press-event", self._on_key_press)
        _style(self, "see-docx-window")

        provider = Gtk.CssProvider()
        provider.load_from_data(_app_css(self))
        screen = Gdk.Screen.get_default()
        if screen:
            Gtk.StyleContext.add_provider_for_screen(
                screen, provider, Gtk.STYLE_PROVIDER_PRIORITY_APPLICATION
            )

        # A grid gives the document an explicit expanding row and keeps the
        # status area in a fixed second row. Gtk.Box otherwise leaves surplus
        # height between start- and end-packed children in this overlay layout.
        root = Gtk.Grid()
        _style(root, "see-docx-root")
        self.document = PdfDocumentView(self._update_pagination_controls)
        workspace = Gtk.Overlay()
        workspace.set_hexpand(True)
        workspace.set_vexpand(True)
        self._document_layout = Gtk.Grid()
        self._document_layout.set_column_homogeneous(False)
        self._document_layout.set_row_homogeneous(False)
        self._document_layout.set_hexpand(True)
        self._document_layout.set_vexpand(True)
        self._outline_panel = self._build_outline_panel()
        self._outline_panel.set_no_show_all(True)
        self._outline_panel.set_halign(Gtk.Align.START)
        self._outline_panel.set_valign(Gtk.Align.FILL)
        self._search_panel = self._build_search_panel()
        self._search_panel.set_no_show_all(True)
        self._search_panel.set_halign(Gtk.Align.CENTER)
        self._search_panel.set_valign(Gtk.Align.CENTER)
        self._page_jump_panel = self._build_page_jump_panel()
        self._page_jump_panel.set_no_show_all(True)
        self._page_jump_panel.set_halign(Gtk.Align.CENTER)
        self._page_jump_panel.set_valign(Gtk.Align.END)
        self._page_jump_panel.set_margin_bottom(18)
        self._export_panel = self._build_export_panel()
        self._export_panel.set_no_show_all(True)
        self._export_panel.set_halign(Gtk.Align.CENTER)
        self._export_panel.set_valign(Gtk.Align.CENTER)
        self._comments_panel = self._build_comments_panel()
        self._comments_revealer = Gtk.Revealer()
        self._comments_revealer.set_transition_type(Gtk.RevealerTransitionType.SLIDE_RIGHT)
        self._comments_revealer.set_transition_duration(180)
        self._comments_revealer.set_reveal_child(False)
        self._comments_revealer.set_hexpand(False)
        self._comments_revealer.set_halign(Gtk.Align.FILL)
        self._comments_revealer.add(self._comments_panel)
        self._comments_revealer.connect(
            "size-allocate", lambda *_args: self._layout_comments()
        )
        self._url_hint_layer = Gtk.Fixed()
        self._url_hint_layer.set_halign(Gtk.Align.FILL)
        self._url_hint_layer.set_valign(Gtk.Align.FILL)
        self._url_hint_layer.set_hexpand(True)
        self._url_hint_layer.set_vexpand(True)
        document_canvas = Gtk.Overlay()
        document_canvas.set_hexpand(True)
        document_canvas.set_vexpand(True)
        self.document.widget.set_hexpand(True)
        self.document.widget.set_vexpand(True)
        self.document.widget.connect("focus-in-event", self._on_document_focus_in)
        document_canvas.add(self.document.widget)
        document_canvas.add_overlay(self._search_panel)
        document_canvas.add_overlay(self._export_panel)
        self._document_layout.attach(self._outline_panel, 0, 0, 1, 1)
        self._document_layout.attach(document_canvas, 1, 0, 1, 1)
        self._document_layout.attach(self._comments_revealer, 2, 0, 1, 1)
        workspace.add(self._document_layout)
        self._comment_line_layer = Gtk.DrawingArea()
        self._comment_line_layer.set_halign(Gtk.Align.FILL)
        self._comment_line_layer.set_valign(Gtk.Align.FILL)
        self._comment_line_layer.set_hexpand(True)
        self._comment_line_layer.set_vexpand(True)
        self._comment_line_layer.set_can_focus(False)
        self._comment_line_layer.set_visible(False)
        self._comment_line_layer.connect("draw", self._draw_comment_connectors)
        self._document_layout.connect(
            "size-allocate", lambda *_args: self._on_document_layout_allocate()
        )
        self.document.widget.connect(
            "size-allocate", lambda *_args: self._layout_comments()
        )
        workspace.add_overlay(self._comment_line_layer)
        workspace.set_overlay_pass_through(self._comment_line_layer, True)
        workspace.add_overlay(self._page_jump_panel)
        workspace.add_overlay(self._url_hint_layer)
        workspace.set_overlay_pass_through(self._url_hint_layer, True)
        status_bar = _FooterStatusBar()
        _style(status_bar, "see-docx-status")
        status_content = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=12)
        status_bar.add(status_content)
        self._path_status = _label(_compact_path(self.path))
        self._path_status.set_hexpand(True)
        self._path_status.set_ellipsize(Pango.EllipsizeMode.MIDDLE)
        self._path_status.set_max_width_chars(72)
        _style(self._path_status, "see-docx-path-status")
        status_content.pack_start(self._path_status, True, True, 0)
        self._page_indicator = _label("Page —", xalign=1.0)
        self._page_indicator.set_width_chars(11)
        _style(self._page_indicator, "see-docx-page-indicator")
        status_content.pack_end(self._page_indicator, False, False, 0)
        self._search_session_status = _label("", xalign=0.5)
        self._search_session_status.set_no_show_all(True)
        self._search_session_status.hide()
        _style(self._search_session_status, "see-docx-search-session")
        status_bar.add_centred_overlay(self._search_session_status)
        self._search_session_status.set_halign(Gtk.Align.CENTER)
        self._search_session_status.set_valign(Gtk.Align.CENTER)
        self._reading_progress = Gtk.ProgressBar()
        self._reading_progress.set_show_text(False)
        self._reading_progress.set_can_focus(False)
        self._reading_progress.set_fraction(0.0)
        self._reading_progress.set_size_request(-1, 3)
        _style(self._reading_progress, "see-docx-reading-progress")
        self._search_match_marker_layer = Gtk.DrawingArea()
        self._search_match_marker_layer.set_halign(Gtk.Align.FILL)
        self._search_match_marker_layer.set_valign(Gtk.Align.FILL)
        self._search_match_marker_layer.set_hexpand(True)
        self._search_match_marker_layer.set_vexpand(False)
        self._search_match_marker_layer.set_can_focus(False)
        self._search_match_marker_layer.connect(
            "draw", self._draw_search_match_markers
        )
        progress_rule = Gtk.Overlay()
        progress_rule.set_size_request(-1, 3)
        progress_rule.set_vexpand(False)
        progress_rule.add(self._reading_progress)
        progress_rule.add_overlay(self._search_match_marker_layer)
        status_area = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        status_area.set_vexpand(False)
        status_area.pack_start(progress_rule, False, False, 0)
        status_area.pack_start(status_bar, False, False, 0)
        root.attach(workspace, 0, 0, 1, 1)
        root.attach(status_area, 0, 1, 1, 1)
        self.add(root)

        self._watch_source()
        self._queue_refresh(delay=0)

    def _build_outline_panel(self) -> Gtk.Box:
        """Build the hidden document-heading navigator."""

        panel = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        panel.set_size_request(280, -1)
        _style(panel, "see-docx-outline")

        title = _label("DOCUMENT STRUCTURE")
        _style(title, "see-docx-outline-title")
        panel.pack_start(title, False, False, 0)

        self._outline_store = Gtk.TreeStore(str, str)
        self._outline_tree = Gtk.TreeView(model=self._outline_store)
        self._outline_tree.set_margin_start(OUTLINE_CONTENT_MARGIN_START)
        self._outline_tree.set_headers_visible(False)
        self._outline_tree.set_enable_search(False)
        self._outline_tree.set_tooltip_text(
            "Offsets show j/k distance · Ctrl-d/u half page · 4j move four · Enter jump · Tab close"
        )
        reference_renderer = Gtk.CellRendererText()
        reference_renderer.set_property("foreground", _theme_palette(self)["muted"])
        reference_renderer.set_property("scale", 0.78)
        reference_renderer.set_property("xalign", 1.0)
        reference_renderer.set_property("xpad", 0)
        reference_renderer.set_property("width-chars", 3)
        title_renderer = Gtk.CellRendererText()
        title_renderer.set_property("ellipsize", Pango.EllipsizeMode.END)
        title_renderer.set_property("width-chars", OUTLINE_HEADING_MAX_CHARS)
        reference_column = Gtk.TreeViewColumn("Relative offset")
        reference_column.pack_start(reference_renderer, True)
        reference_column.add_attribute(reference_renderer, "text", 0)
        reference_column.set_sizing(Gtk.TreeViewColumnSizing.FIXED)
        reference_column.set_fixed_width(OUTLINE_REFERENCE_MARGIN_WIDTH)
        title_column = Gtk.TreeViewColumn("Heading")
        title_column.pack_start(title_renderer, True)
        title_column.add_attribute(title_renderer, "text", 1)
        self._outline_tree.append_column(reference_column)
        self._outline_tree.append_column(title_column)
        # Keep the relative movement offsets in a stable margin. The expander
        # and title share their content column, so tree nesting never inserts
        # a number between an arrow and its heading.
        self._outline_tree.set_expander_column(title_column)
        self._outline_tree.connect("key-press-event", self._on_key_press)
        self._outline_tree.connect("row-activated", self._on_outline_row_activated)
        self._outline_tree.get_selection().connect(
            "changed", self._on_outline_selection_changed
        )
        self._outline_tree.connect(
            "row-expanded", self._on_outline_row_visibility_changed
        )
        self._outline_tree.connect(
            "row-collapsed", self._on_outline_row_visibility_changed
        )

        scroller = Gtk.ScrolledWindow()
        scroller.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        scroller.add(self._outline_tree)
        panel.pack_start(scroller, True, True, 0)

        self._outline_empty = _label("No headings found in this document.")
        self._outline_empty.set_line_wrap(True)
        _style(self._outline_empty, "see-docx-outline-empty")
        panel.pack_start(self._outline_empty, False, False, 0)
        return panel

    def _build_comments_panel(self) -> Gtk.Box:
        """Build the responsive, independently scrollable comment rail."""

        panel = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        panel.set_size_request(COMMENTS_PANEL_MIN_WIDTH, -1)
        panel.set_hexpand(False)
        _style(panel, "see-docx-comments")
        panel.set_tooltip_text(
            "Comments · c focus · j/k select · Enter read · v show or hide"
        )

        header = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=2)
        _style(header, "see-docx-comments-header")
        header_top = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=8)
        _style(header_top, "see-docx-comments-header-top")
        title = _label("COMMENTS")
        _style(title, "see-docx-comments-title")
        header_top.pack_start(title, True, True, 0)
        self._comments_count = _label("0 / 0", xalign=1.0)
        _style(self._comments_count, "see-docx-comments-count")
        header_top.pack_end(self._comments_count, False, False, 0)
        header.pack_start(header_top, False, False, 0)

        header_meta = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=7)
        _style(header_meta, "see-docx-comments-header-meta")
        self._comments_summary = _label("THREADS", xalign=0.0)
        _style(self._comments_summary, "see-docx-comments-summary")
        header_meta.pack_start(self._comments_summary, True, True, 0)
        toggle_hint = _label("v hide", xalign=1.0)
        _style(toggle_hint, "see-docx-comments-key")
        header_meta.pack_end(toggle_hint, False, False, 0)
        focus_hint = _label("c focus", xalign=1.0)
        _style(focus_hint, "see-docx-comments-key")
        header_meta.pack_end(focus_hint, False, False, 0)
        header.pack_start(header_meta, False, False, 0)
        panel.pack_start(header, False, False, 0)

        self._comments_list = Gtk.Box(
            orientation=Gtk.Orientation.VERTICAL, spacing=COMMENT_CARD_GAP
        )
        self._comments_list.set_hexpand(True)
        self._comments_list.set_vexpand(True)
        self._comments_list.set_margin_start(COMMENTS_PANEL_GUTTER)
        self._comments_list.set_margin_end(COMMENTS_PANEL_GUTTER)
        self._comments_list.set_margin_top(COMMENT_CARD_GAP)
        self._comments_list.set_margin_bottom(COMMENT_CARD_GAP)
        self._comments_empty = _label("No comments in this document.")
        self._comments_empty.set_line_wrap(True)
        self._comments_empty.set_max_width_chars(26)
        _style(self._comments_empty, "see-docx-comments-empty")
        self._comments_list.pack_start(self._comments_empty, False, False, 0)
        self._comments_scroller = Gtk.ScrolledWindow()
        self._comments_scroller.set_policy(
            Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC
        )
        self._comments_scroller.set_kinetic_scrolling(True)
        self._comments_scroller.set_hexpand(True)
        self._comments_scroller.set_vexpand(True)
        self._comments_scroller.add(self._comments_list)
        panel.pack_start(self._comments_scroller, True, True, 0)
        return panel

    @staticmethod
    def _comment_date_label(value: str | None) -> str:
        if not value:
            return ""
        try:
            return datetime.fromisoformat(value.replace("Z", "+00:00")).strftime(
                "%d %b"
            )
        except ValueError:
            return ""

    def _build_comment_message(
        self, comment: DocumentComment, *, reply: bool
    ) -> Gtk.Box:
        """Build one message inside a grouped conversation."""

        message = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=5)
        _style(message, "see-docx-comment-message")
        if reply:
            _style(message, "reply")
        metadata = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=7)
        if reply:
            reply_label = _label("↳ REPLY")
            _style(reply_label, "see-docx-comment-reply-label")
            metadata.pack_start(reply_label, False, False, 0)
        initials = _label(comment.initials or "?", xalign=0.5)
        _style(initials, "see-docx-comment-initials")
        metadata.pack_start(initials, False, False, 0)
        author = _label(comment.author or "Comment")
        _style(author, "see-docx-comment-author")
        metadata.pack_start(author, False, False, 0)
        date = self._comment_date_label(comment.date)
        if date:
            date_label = _label(date, xalign=1.0)
            _style(date_label, "see-docx-comment-date")
            metadata.pack_end(date_label, False, False, 0)
        message.pack_start(metadata, False, False, 0)

        body = _label(comment.text or "No comment text.")
        body.set_line_wrap(True)
        body.set_line_wrap_mode(Pango.WrapMode.WORD_CHAR)
        body.set_xalign(0.0)
        body.set_yalign(0.0)
        _style(body, "see-docx-comment-body")
        message.pack_start(body, False, False, 0)
        return message

    def _build_comment_card(self, thread: CommentThread) -> Gtk.Box:
        """Create one selectable conversation with a single scroll surface."""

        bubble = Gtk.Box(orientation=Gtk.Orientation.VERTICAL, spacing=7)
        bubble.set_hexpand(True)
        bubble.set_can_focus(True)
        bubble.add_events(Gdk.EventMask.BUTTON_PRESS_MASK)
        bubble.connect(
            "button-press-event", self._on_comment_card_press, thread.thread_id
        )
        _style(bubble, "see-docx-comment-thread")
        bubble.set_halign(Gtk.Align.FILL)

        if thread.reply_count:
            thread_label = _label(
                f"THREAD · {len(thread.comments)} MESSAGES",
                xalign=0.0,
            )
            _style(thread_label, "see-docx-comment-thread-label")
            bubble.pack_start(thread_label, False, False, 0)

        conversation = Gtk.Box(
            orientation=Gtk.Orientation.VERTICAL,
            spacing=8,
        )
        conversation.pack_start(
            self._build_comment_message(thread.root, reply=False),
            False,
            False,
            0,
        )
        for reply in thread.comments[1:]:
            conversation.pack_start(
                self._build_comment_message(reply, reply=True),
                False,
                False,
                0,
            )

        body_scroller = Gtk.ScrolledWindow()
        body_scroller.set_policy(Gtk.PolicyType.NEVER, Gtk.PolicyType.AUTOMATIC)
        body_scroller.set_kinetic_scrolling(True)
        body_scroller.set_hexpand(True)
        body_scroller.set_can_focus(True)
        _style(body_scroller, "see-docx-comment-body-scroll")
        body_scroller.add(conversation)
        body_scroller.connect(
            "size-allocate",
            lambda *_args: GLib.idle_add(self._layout_comments_after_allocate),
        )
        body_scroller.connect("focus-in-event", self._on_comment_focus_in)
        bubble.pack_start(body_scroller, False, False, 0)
        self._comment_body_scrollers[thread.thread_id] = body_scroller

        anchor = self.document.comment_anchor_text(thread.root).strip()
        if anchor:
            anchor = " ".join(anchor.split())
            if len(anchor) > 96:
                anchor = f"{anchor[:93].rstrip()}…"
            anchor_label = _label(f'“{anchor}”')
            anchor_label.set_line_wrap(True)
            anchor_label.set_line_wrap_mode(Pango.WrapMode.WORD_CHAR)
            anchor_label.set_xalign(0.0)
            _style(anchor_label, "see-docx-comment-anchor")
            bubble.pack_start(anchor_label, False, False, 0)
        return bubble

    def _on_comment_focus_in(
        self, _widget: Gtk.Widget, _event: Gdk.EventFocus
    ) -> bool:
        """Mark the rail and body levels when a comment body gains focus."""

        self._comments_focused = True
        self._comment_body_focused = True
        self._comments_panel.get_style_context().add_class("comments-focused")
        self._comments_panel.get_style_context().add_class("comments-body-focused")
        self._update_comment_count()
        self._sync_active_comment_mark()
        self._apply_comment_sizing()
        return False

    def _on_document_focus_in(
        self, _widget: Gtk.Widget, _event: Gdk.EventFocus
    ) -> bool:
        """Drop the rail highlight when the document regains focus."""

        self._comments_focused = False
        self._comment_body_focused = False
        self._comments_panel.get_style_context().remove_class("comments-focused")
        self._comments_panel.get_style_context().remove_class(
            "comments-body-focused"
        )
        self._update_comment_count()
        self._sync_active_comment_mark()
        self._apply_comment_sizing()
        return False

    def _update_comment_count(self) -> None:
        """Show a selection index only while the comment rail is focused."""

        total = len(self._comment_annotations)
        if self._comments_focused and self._active_comment_index >= 0:
            self._comments_count.set_text(
                f"{self._active_comment_index + 1} / {total}"
            )
        else:
            self._comments_count.set_text(f"— / {total}")

    def _sync_active_comment_mark(self) -> None:
        """Expose the focused comment to the document mark renderer."""

        self.document.set_active_comment(
            self._active_comment_id if self._comments_focused else None
        )

    def _focus_comments(self) -> None:
        """Enter the comment rail without changing the active comment."""

        if not self._comment_annotations:
            return
        self._pending_g = False
        if not self._comments_visible:
            self._toggle_comments()
        self._comments_focused = True
        self._comment_body_focused = False
        self._comments_panel.get_style_context().add_class("comments-focused")
        self._comments_panel.get_style_context().remove_class(
            "comments-body-focused"
        )
        self._update_comment_count()
        self._sync_active_comment_mark()
        self._apply_comment_sizing()
        GLib.idle_add(self._scroll_active_comment_into_view)
        GLib.idle_add(self._focus_active_comment_card)

    def _blur_comments(self) -> None:
        """Leave the rail and return keyboard navigation to the document."""

        self._comments_focused = False
        self._pending_g = False
        self._comment_body_focused = False
        self._comments_panel.get_style_context().remove_class("comments-focused")
        self._comments_panel.get_style_context().remove_class(
            "comments-body-focused"
        )
        self._update_comment_count()
        self._sync_active_comment_mark()
        self._apply_comment_sizing()
        self.document.widget.grab_focus()

    def _leave_comment_body(self) -> None:
        """Return from comment text scrolling to the comment list."""

        if not self._comments_focused:
            return
        self._pending_g = False
        self._comment_body_focused = False
        self._comments_panel.get_style_context().remove_class(
            "comments-body-focused"
        )
        self._focus_active_comment_card()

    def _set_comments(self, comments: Iterable[DocumentComment]) -> None:
        """Replace the rail contents and reveal it only when threads exist."""

        was_available = self._comments_available
        for card in self._comment_cards.values():
            card.destroy()
        self._comment_cards.clear()
        self._comment_body_scrollers.clear()
        self._comment_body_focused = False
        self._comments_panel.get_style_context().remove_class(
            "comments-body-focused"
        )
        self._comment_annotations = _comment_threads(tuple(comments))
        has_comments = bool(self._comment_annotations)
        self._comments_available = has_comments
        if not has_comments:
            self._comments_visible = False
            self._comments_focused = False
            self._comments_panel.get_style_context().remove_class(
                "comments-focused"
            )
        elif not was_available:
            # First content load, or a refresh from an unannotated document:
            # comments are a useful secondary surface by default.
            self._comments_visible = True
        self._comments_revealer.set_reveal_child(self._comments_visible)
        self._comment_line_layer.set_visible(self._comments_visible)
        anchors_by_comment_id = {
            anchor.comment.comment_id: anchor
            for anchor in self.document.comment_anchors
        }
        self._comment_anchor_by_id = {}
        for thread in self._comment_annotations:
            anchor = anchors_by_comment_id.get(thread.thread_id)
            if anchor is None:
                anchor = next(
                    (
                        anchors_by_comment_id.get(comment.comment_id)
                        for comment in thread.comments
                        if comment.comment_id in anchors_by_comment_id
                    ),
                    None,
                )
            if anchor is not None:
                self._comment_anchor_by_id[thread.thread_id] = anchor
        self._active_comment_index = 0 if self._comment_annotations else -1
        self._active_comment_id = (
            self._comment_annotations[0].comment_id
            if self._comment_annotations
            else None
        )
        message_count = sum(
            len(thread.comments) for thread in self._comment_annotations
        )
        self._comments_summary.set_text(
            f"{len(self._comment_annotations)} threads · {message_count} messages"
            if message_count
            else "THREADS"
        )
        self._update_comment_count()
        self._sync_active_comment_mark()
        self._comments_empty.set_visible(not self._comment_annotations)
        for thread in self._comment_annotations:
            card = self._build_comment_card(thread)
            self._comment_cards[thread.thread_id] = card
            self._comments_list.pack_start(card, False, False, 0)
            card.show_all()
        self._apply_comment_sizing()

    def _apply_comment_sizing(self) -> None:
        """Apply the active width instantly, growing only toward the left."""

        for thread in self._comment_annotations:
            card = self._comment_cards.get(thread.thread_id)
            if card is None:
                continue
            active = (
                self._comments_focused
                and thread.thread_id == self._active_comment_id
            )
            card.set_margin_start(
                0
                if active
                else COMMENT_INACTIVE_CARD_INSET
            )
            card.set_margin_end(0)
        self._layout_comments()

    def _on_comment_card_press(
        self, _card: Gtk.Box, event: Gdk.EventButton, thread_id: str
    ) -> bool:
        """Make a pointer-selected card the active connector target."""

        if event.button != 1:
            return False
        self._focus_comments()
        for index, thread in enumerate(self._comment_annotations):
            if thread.thread_id == thread_id:
                self._activate_comment(index, reveal_document=False)
                break
        return False

    def _activate_comment(self, index: int, *, reveal_document: bool) -> None:
        """Select one comment and keep its card and anchor in view."""

        if not self._comment_annotations:
            self._active_comment_index = -1
            self._active_comment_id = None
            return
        index = min(max(index, 0), len(self._comment_annotations) - 1)
        thread = self._comment_annotations[index]
        self._active_comment_index = index
        self._active_comment_id = thread.thread_id
        self._update_comment_count()
        self._sync_active_comment_mark()
        if reveal_document:
            anchor = self._comment_anchor_by_id.get(thread.thread_id)
            if anchor is not None:
                self.document.go_to_comment_anchor(anchor)
        self._apply_comment_sizing()
        GLib.idle_add(self._scroll_active_comment_into_view)
        if self._comments_focused and not self._comment_body_focused:
            GLib.idle_add(self._focus_active_comment_card)

    def _focus_active_comment_card(self) -> bool:
        """Keep keyboard focus on the selected card in the comment list."""

        if (
            not self._comments_focused
            or self._comment_body_focused
            or self._active_comment_id is None
        ):
            return GLib.SOURCE_REMOVE
        card = self._comment_cards.get(self._active_comment_id)
        if card is not None:
            card.grab_focus()
        return GLib.SOURCE_REMOVE

    def _enter_comment_body(self) -> None:
        """Enter text scrolling for the currently selected comment."""

        if not self._comments_focused or self._active_comment_id is None:
            return
        body = self._comment_body_scrollers.get(self._active_comment_id)
        if body is None:
            return
        self._comment_body_focused = True
        self._comments_panel.get_style_context().add_class(
            "comments-body-focused"
        )
        body.grab_focus()

    def _move_comment_selection(self, direction: int) -> None:
        """Move the selected thread, stopping at either end."""

        if not self._comment_annotations:
            return
        if not self._comments_visible:
            self._toggle_comments()
        if not self._comments_focused:
            self._focus_comments()
            return
        if self._comment_body_focused:
            return
        index = self._active_comment_index
        if index < 0:
            index = 0 if direction > 0 else len(self._comment_annotations) - 1
        else:
            index = min(
                max(index + direction, 0),
                len(self._comment_annotations) - 1,
            )
        self._activate_comment(index, reveal_document=True)

    def _scroll_active_comment_into_view(self) -> bool:
        """Center the active card without coupling rail scroll to page scroll."""

        if not self._comments_visible or self._active_comment_id is None:
            return GLib.SOURCE_REMOVE
        card = self._comment_cards.get(self._active_comment_id)
        if card is None or card.get_allocated_height() <= 0:
            return GLib.SOURCE_REMOVE
        adjustment = self._comments_scroller.get_vadjustment()
        page_size = adjustment.get_page_size()
        top = float(card.get_allocation().y)
        bottom = top + float(card.get_allocated_height())
        target = (top + bottom) / 2 - page_size / 2
        maximum = max(
            adjustment.get_lower(),
            adjustment.get_upper() - adjustment.get_page_size(),
        )
        adjustment.set_value(min(max(target, adjustment.get_lower()), maximum))
        self._comment_line_layer.queue_draw()
        return GLib.SOURCE_REMOVE

    def _active_comment_body(self) -> Gtk.ScrolledWindow | None:
        """Return the active body when keyboard focus is inside its scroller."""

        if (
            not self._comments_focused
            or not self._comment_body_focused
            or self._active_comment_id is None
        ):
            return None
        body = self._comment_body_scrollers.get(self._active_comment_id)
        if body is None or not body.is_focus():
            return None
        return body

    @staticmethod
    def _scroll_comment_body(body: Gtk.ScrolledWindow, direction: int) -> None:
        """Scroll an active comment body by one line or half a viewport."""

        adjustment = body.get_vadjustment()
        if direction in {-1, 1}:
            amount = SCROLL_STEP
        else:
            amount = adjustment.get_page_size() / 2
            direction = 1 if direction > 0 else -1
        lower = adjustment.get_lower()
        maximum = max(lower, adjustment.get_upper() - adjustment.get_page_size())
        adjustment.set_value(
            min(max(adjustment.get_value() + direction * amount, lower), maximum)
        )

    def _scroll_comment_list(self, direction: int) -> None:
        """Scroll the full comment list by half a viewport."""

        adjustment = self._comments_scroller.get_vadjustment()
        amount = adjustment.get_page_size() / 2
        lower = adjustment.get_lower()
        maximum = max(lower, adjustment.get_upper() - adjustment.get_page_size())
        adjustment.set_value(
            min(max(adjustment.get_value() + direction * amount, lower), maximum)
        )

    @staticmethod
    def _translated_point(
        widget: Gtk.Widget,
        target: Gtk.Widget,
        x: float,
        y: float,
    ) -> tuple[float, float] | None:
        translated = widget.translate_coordinates(target, round(x), round(y))
        if translated is None:
            return None
        if len(translated) == 3:
            success, translated_x, translated_y = translated
            if not success:
                return None
            return float(translated_x), float(translated_y)
        if len(translated) == 2:
            return float(translated[0]), float(translated[1])
        return None

    def _layout_comments_after_allocate(self) -> bool:
        """Retry once GTK has committed the page stack's final coordinates."""

        self._layout_comments()
        return GLib.SOURCE_REMOVE

    def _layout_comments(self) -> None:
        """Refresh responsive card sizing and the active connector."""

        if not hasattr(self, "_comment_line_layer") or not self._comments_visible:
            return
        for thread in self._comment_annotations:
            body_scroller = self._comment_body_scrollers.get(thread.thread_id)
            card = self._comment_cards.get(thread.thread_id)
            if body_scroller is None or card is None:
                continue
            active = (
                self._comments_focused
                and thread.thread_id == self._active_comment_id
            )
            style_context = card.get_style_context()
            if active:
                style_context.add_class("active")
            else:
                style_context.remove_class("active")
            body = body_scroller.get_child()
            width = max(
                1,
                body_scroller.get_allocated_width()
                or self._comments_panel.get_allocated_width()
                - 2 * COMMENTS_PANEL_GUTTER
                - 24,
            )
            natural_height = 0
            if body is not None and hasattr(body, "get_preferred_height_for_width"):
                _minimum, natural_height = body.get_preferred_height_for_width(width)
            content_height = max(COMMENT_BODY_MIN_HEIGHT, natural_height + 4)
            height = min(
                COMMENT_ACTIVE_BODY_MAX_HEIGHT
                if active
                else COMMENT_INACTIVE_BODY_HEIGHT,
                content_height,
            )
            body_scroller.set_size_request(-1, max(COMMENT_BODY_MIN_HEIGHT, height))
        self._comments_empty.set_visible(not self._comment_annotations)
        self._comment_line_layer.queue_draw()

    def _on_document_layout_allocate(self) -> None:
        """Give the rail more width on wide windows without starving the page."""

        available = self._document_layout.get_allocated_width()
        if available <= 0:
            return
        width = min(
            COMMENTS_PANEL_MAX_WIDTH,
            max(COMMENTS_PANEL_MIN_WIDTH, round(available * COMMENTS_PANEL_FRACTION)),
        )
        if available - width < COMMENTS_DOCUMENT_MIN_WIDTH:
            width = max(240, available - COMMENTS_DOCUMENT_MIN_WIDTH)
        current = self._comments_panel.get_allocated_width()
        requested = getattr(self, "_comment_panel_width", 0)
        if requested != width:
            self._comment_panel_width = width
            self._comments_panel.set_size_request(width, -1)
        if current != width:
            GLib.idle_add(self._layout_comments_after_allocate)

    def _draw_comment_connectors(
        self, _layer: Gtk.DrawingArea, context: object
    ) -> bool:
        """Draw the thin, paper-to-rail tethers behind the comment bubbles."""

        if not self._comments_visible:
            return False
        palette = _theme_palette(self)
        accent = _rgb(palette["accent"])
        layer_width = self._comment_line_layer.get_allocated_width()
        layer_height = self._comment_line_layer.get_allocated_height()
        if layer_width <= 0 or layer_height <= 0:
            return False
        active_id = self._active_comment_id if self._comments_focused else None
        for thread in self._comment_annotations:
            if thread.thread_id != active_id:
                continue
            card = self._comment_cards.get(thread.thread_id)
            anchor = self._comment_anchor_by_id.get(thread.thread_id)
            if card is None or anchor is None or not card.get_visible():
                continue
            if not 0 <= anchor.page_index < len(self.document._pages):
                continue
            page = self.document._pages[anchor.page_index]
            start = self._translated_point(
                page,
                self._comment_line_layer,
                anchor.right * page._zoom + 3,
                anchor.center_y * page._zoom,
            )
            card_point = self._translated_point(
                card,
                self._comment_line_layer,
                0,
                card.get_allocated_height() / 2,
            )
            if start is None or card_point is None:
                continue
            start_x, start_y = start
            card_x, card_y = card_point
            if not (
                -COMMENT_VIEWPORT_BUFFER <= start_y <= layer_height + COMMENT_VIEWPORT_BUFFER
                and -COMMENT_VIEWPORT_BUFFER <= card_y <= layer_height + COMMENT_VIEWPORT_BUFFER
            ):
                continue
            page_right = self._translated_point(
                page,
                self._comment_line_layer,
                page.get_allocated_width(),
                anchor.center_y * page._zoom,
            )
            if page_right is None:
                continue
            card_edge = card_x - 8
            if start_x >= card_edge:
                continue
            edge_x = min(page_right[0] + 10, card_edge - 12)
            edge_x = max(edge_x, start_x)
            context.save()
            context.set_line_cap(cairo.LINE_CAP_ROUND)
            context.set_line_join(cairo.LINE_JOIN_ROUND)
            context.set_source_rgba(*accent, 0.16)
            context.set_line_width(4.0)
            context.move_to(start_x, start_y)
            context.line_to(edge_x, start_y)
            context.line_to(card_edge, start_y)
            context.line_to(card_edge, card_y)
            context.stroke()
            context.set_source_rgba(*accent, 0.72)
            context.set_line_width(1.25)
            context.move_to(start_x, start_y)
            context.line_to(edge_x, start_y)
            context.line_to(card_edge, start_y)
            context.line_to(card_edge, card_y)
            context.stroke()
            context.arc(start_x, start_y, 3.0, 0, 2 * math.pi)
            context.fill()
            context.restore()
        return False

    def _build_search_panel(self) -> Gtk.Box:
        """Build the centered, single-surface in-document search prompt."""

        panel = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        _style(panel, "see-docx-search")
        panel.set_size_request(380, -1)
        prefix = _label("/")
        _style(prefix, "see-docx-search-prefix")
        panel.pack_start(prefix, False, False, 0)
        self._search_entry = Gtk.Entry()
        self._search_entry.set_placeholder_text("Search document")
        self._search_entry.set_width_chars(32)
        self._search_entry.set_hexpand(True)
        self._search_entry.connect("changed", self._on_search_changed)
        self._search_entry.connect("key-press-event", self._on_search_key_press)
        panel.pack_start(self._search_entry, True, True, 0)
        self._search_status = _label("")
        _style(self._search_status, "see-docx-search-status")
        panel.pack_start(self._search_status, False, False, 0)
        return panel

    def _build_page_jump_panel(self) -> Gtk.Box:
        """Build a compact one-based page-jump prompt."""

        panel = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL)
        _style(panel, "see-docx-page-jump")
        prefix = _label(":")
        _style(prefix, "see-docx-page-jump-prefix")
        panel.pack_start(prefix, False, False, 0)
        self._page_jump_entry = Gtk.Entry()
        self._page_jump_entry.set_placeholder_text("Page number")
        self._page_jump_entry.set_width_chars(10)
        self._page_jump_entry.connect("key-press-event", self._on_page_jump_key_press)
        panel.pack_start(self._page_jump_entry, True, True, 0)
        self._page_jump_status = _label("")
        _style(self._page_jump_status, "see-docx-page-jump-status")
        panel.pack_start(self._page_jump_status, False, False, 0)
        return panel

    def _build_export_panel(self) -> Gtk.Box:
        """Build the keyboard-first document export chooser."""

        panel = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        _style(panel, "see-docx-export")
        title = _label("EXPORT AS")
        _style(title, "see-docx-export-title")
        panel.pack_start(title, False, False, 0)

        self._export_list = Gtk.ListBox()
        self._export_list.set_selection_mode(Gtk.SelectionMode.SINGLE)
        self._export_list.connect("row-activated", self._on_export_row_activated)
        for export_format in EXPORT_FORMATS:
            row = Gtk.ListBoxRow()
            row_content = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
            format_label = _label(export_format)
            _style(format_label, "see-docx-export-format")
            description = _label(EXPORT_FORMAT_DESCRIPTIONS[export_format])
            _style(description, "see-docx-export-description")
            row_content.pack_start(format_label, False, False, 0)
            row_content.pack_start(description, False, False, 0)
            row.add(row_content)
            self._export_list.add(row)
        panel.pack_start(self._export_list, False, False, 0)

        self._export_status = _label("j/k select · Enter choose destination · Esc close")
        self._export_status.set_line_wrap(True)
        _style(self._export_status, "see-docx-export-status")
        panel.pack_start(self._export_status, False, False, 0)
        return panel

    def _update_outline(self) -> None:
        """Replace the sidebar model after a newly rendered PDF is loaded."""

        self._outline_entries = list(self.document.outline)
        self._outline_row_paths.clear()
        self._outline_store.clear()
        parents: dict[int, Gtk.TreeIter] = {}
        for entry in self._outline_entries:
            parent = parents.get(entry.depth - 1)
            row = self._outline_store.append(parent, ["", entry.title])
            self._outline_row_paths.append(self._outline_store.get_path(row))
            parents[entry.depth] = row
            for child_depth in tuple(parents):
                if child_depth > entry.depth:
                    del parents[child_depth]
        self._set_outline_default_expansion()
        self._outline_empty.set_visible(not self._outline_entries)
        self._outline_tree.set_visible(bool(self._outline_entries))
        if self._outline_panel.get_visible():
            self._select_nearest_outline_entry()

    def _selected_outline_index(self) -> int | None:
        model, row = self._outline_tree.get_selection().get_selected()
        if row is None:
            return None
        path = model.get_path(row)
        for index, candidate in enumerate(self._outline_row_paths):
            if candidate.compare(path) == 0:
                return index
        return None

    def _set_outline_default_expansion(self) -> None:
        """Reveal only the heading depth appropriate to this document's size."""

        maximum_depth = _outline_default_expansion_depth(self._outline_entries)
        self._outline_tree.collapse_all()
        for entry, path in zip(self._outline_entries, self._outline_row_paths):
            if entry.depth < maximum_depth:
                self._outline_tree.expand_row(path, False)

    def _select_outline_entry(self, index: int) -> bool:
        if not 0 <= index < len(self._outline_row_paths):
            return False
        path = self._outline_row_paths[index]
        self._outline_tree.get_selection().select_path(path)
        self._outline_tree.scroll_to_cell(path, None, False, 0.0, 0.0)
        self._update_outline_references()
        return True

    def _visible_outline_indices(self) -> list[int]:
        """Return heading rows whose ancestor branches are currently open."""

        visible: list[int] = []
        for index, path in enumerate(self._outline_row_paths):
            ancestors = (
                ancestor
                for ancestor in self._outline_row_paths[:index]
                if ancestor.is_ancestor(path)
            )
            if all(self._outline_tree.row_expanded(ancestor) for ancestor in ancestors):
                visible.append(index)
        return visible

    def _update_outline_references(self) -> None:
        """Label each visible heading by its movement offset from selection."""

        selected = self._selected_outline_index()
        visible = self._visible_outline_indices()
        if selected is None:
            references: dict[int, str] = {}
        else:
            try:
                selected_position = visible.index(selected)
            except ValueError:
                references = {}
            else:
                references = {
                    index: str(position - selected_position)
                    for position, index in enumerate(visible)
                }
        for index, path in enumerate(self._outline_row_paths):
            row = self._outline_store.get_iter(path)
            reference = references.get(index, "")
            if self._outline_store.get_value(row, 0) != reference:
                self._outline_store.set_value(row, 0, reference)

    def _on_outline_selection_changed(self, _selection: Gtk.TreeSelection) -> None:
        self._update_outline_references()

    def _on_outline_row_visibility_changed(
        self, _tree: Gtk.TreeView, _row: Gtk.TreeIter, _path: Gtk.TreePath
    ) -> None:
        self._update_outline_references()

    def _select_nearest_outline_entry(self) -> bool:
        visible = self._visible_outline_indices()
        if not visible:
            return False
        current_page = self.document.current_page_index or 0
        index = visible[-1]
        for candidate in visible:
            if self._outline_entries[candidate].page_index >= current_page:
                index = candidate
                break
        return self._select_outline_entry(index)

    def _move_outline_selection(self, direction: int) -> bool:
        visible = self._visible_outline_indices()
        if not visible:
            return False
        index = self._selected_outline_index()
        if index is None:
            return self._select_nearest_outline_entry()
        try:
            position = visible.index(index)
        except ValueError:
            position = next(
                (
                    candidate
                    for candidate, visible_index in enumerate(visible)
                    if visible_index >= index
                ),
                len(visible) - 1,
            )
        return self._select_outline_entry(
            visible[min(max(position + direction, 0), len(visible) - 1)]
        )

    def _outline_half_page_step(self) -> int:
        """Return half the number of heading rows visible in the tree."""

        visible = self._visible_outline_indices()
        if not visible:
            return 1
        visible_range = self._outline_tree.get_visible_range()
        if len(visible_range) == 2:
            # PyGObject returns the out parameters directly and omits GTK's
            # C-level success boolean.
            start_path, end_path = visible_range
        elif len(visible_range) == 3:
            has_range, start_path, end_path = visible_range
            if not has_range:
                return 1
        else:
            return 1
        positions = {
            self._outline_row_paths[index].to_string(): position
            for position, index in enumerate(visible)
        }
        start = positions.get(start_path.to_string())
        end = positions.get(end_path.to_string())
        if start is None or end is None:
            return 1
        return max((end - start + 1) // 2, 1)

    def _set_outline_expanded(self, expanded: bool) -> bool:
        """Expand or collapse the selected heading without changing selection."""

        index = self._selected_outline_index()
        if index is None and not self._select_nearest_outline_entry():
            return False
        index = self._selected_outline_index()
        if index is None:
            return False
        path = self._outline_row_paths[index]
        if expanded:
            self._outline_tree.expand_row(path, False)
        else:
            self._outline_tree.collapse_row(path)
        self._update_outline_references()
        return True

    def _activate_selected_outline_entry(self) -> bool:
        index = self._selected_outline_index()
        if index is None and not self._select_nearest_outline_entry():
            return False
        index = self._selected_outline_index()
        if index is None:
            return False
        entry = self._outline_entries[index]
        if not self.document.go_to_outline_entry(entry):
            return False
        return True

    def _on_outline_row_activated(
        self, _tree: Gtk.TreeView, path: Gtk.TreePath, _column: Gtk.TreeViewColumn
    ) -> None:
        self._outline_tree.get_selection().select_path(path)
        self._activate_selected_outline_entry()

    def _toggle_outline(self) -> None:
        if self._outline_panel.get_visible():
            self._outline_panel.hide()
            if self._outline_fit_source:
                GLib.source_remove(self._outline_fit_source)
                self._outline_fit_source = 0
            if self._outline_zoom_before_open is not None:
                self.document.set_zoom(self._outline_zoom_before_open)
                self._outline_zoom_before_open = None
            return
        self._outline_zoom_before_open = self.document.zoom
        # ``no-show-all`` keeps the initially hidden sidebar from appearing
        # during the window's first show_all(). Lift it just long enough for
        # the title and tree descendants to be realized too.
        self._outline_panel.set_no_show_all(False)
        self._outline_panel.show_all()
        self._outline_panel.set_no_show_all(True)
        self._outline_empty.set_visible(not self._outline_entries)
        self._outline_tree.set_visible(bool(self._outline_entries))
        self._set_outline_default_expansion()
        self._select_nearest_outline_entry()
        self._outline_tree.grab_focus()
        self._outline_fit_source = GLib.idle_add(self._fit_document_to_outline)

    def _toggle_comments(self) -> None:
        """Slide the annotation rail away while leaving the document intact."""

        if not getattr(self, "_comments_available", bool(self._comment_annotations)):
            return
        self._comments_visible = not self._comments_visible
        self._comments_revealer.set_reveal_child(self._comments_visible)
        self._comment_line_layer.set_visible(self._comments_visible)
        if self._comments_visible:
            self._apply_comment_sizing()
            if self._comments_focused:
                GLib.idle_add(self._scroll_active_comment_into_view)
        else:
            self._blur_comments()

    def _fit_document_to_outline(self) -> bool:
        """Fit the PDF after GTK assigns the document column its new width."""

        self._outline_fit_source = 0
        if self._outline_panel.get_visible():
            self.document.fit_to_viewport()
        return GLib.SOURCE_REMOVE

    def _toggle_search(self) -> None:
        """Open or dismiss the transient search prompt without changing content."""

        if self._search_panel.get_visible():
            self._search_panel.hide()
            self.document.widget.grab_focus()
            return
        if self._page_jump_panel.get_visible():
            self._page_jump_panel.hide()
        self._search_panel.set_no_show_all(False)
        self._search_panel.show_all()
        self._search_panel.set_no_show_all(True)
        self._search_entry.grab_focus()
        self._search_entry.select_region(0, -1)

    def _cancel_search(self) -> None:
        """End the current search session and remove its PDF highlight."""

        self._search_session_committed = False
        self._clear_search_session_status()
        self._search_entry.set_text("")
        self._search_matches.clear()
        self._search_index = -1
        self._search_status.set_text("")
        self.document.clear_search_highlight()
        self._queue_search_match_marker_redraw()
        if self._search_panel.get_visible():
            self._search_panel.hide()
        self.document.widget.grab_focus()

    def _toggle_page_jump(self) -> None:
        """Open or dismiss the transient one-based page-number prompt."""

        if self._page_jump_panel.get_visible():
            self._page_jump_panel.hide()
            self.document.widget.grab_focus()
            return
        if self._search_panel.get_visible():
            self._search_panel.hide()
        self._page_jump_entry.set_text("")
        self._page_jump_status.set_text(f"1–{self.document.page_count}")
        self._page_jump_panel.set_no_show_all(False)
        self._page_jump_panel.show_all()
        self._page_jump_panel.set_no_show_all(True)
        self._page_jump_entry.grab_focus()

    def _toggle_export(self) -> None:
        """Open or dismiss the export format chooser."""

        if self._export_panel.get_visible():
            self._export_panel.hide()
            self.document.widget.grab_focus()
            return
        if self._search_panel.get_visible():
            self._search_panel.hide()
        if self._page_jump_panel.get_visible():
            self._page_jump_panel.hide()
        self._export_index = 0
        self._export_list.select_row(self._export_list.get_row_at_index(0))
        self._export_status.set_text("j/k select · Enter choose destination · Esc close")
        self._export_panel.set_no_show_all(False)
        self._export_panel.show_all()
        self._export_panel.set_no_show_all(True)

    def _move_export_selection(self, direction: int) -> None:
        """Move the active export target, wrapping for future formats."""

        self._export_index = (self._export_index + direction) % len(EXPORT_FORMATS)
        self._export_list.select_row(
            self._export_list.get_row_at_index(self._export_index)
        )

    def _on_export_row_activated(
        self, _list: Gtk.ListBox, row: Gtk.ListBoxRow
    ) -> None:
        self._export_index = row.get_index()
        self._activate_export_option()

    def _activate_export_option(self) -> None:
        """Prompt for a destination and start the selected format conversion."""

        if self._export_process is not None:
            self._export_status.set_text("An export is already in progress.")
            return
        if EXPORT_FORMATS[self._export_index] == "PDF":
            self._choose_pdf_destination()
        elif EXPORT_FORMATS[self._export_index] == "Plain text":
            self._choose_plain_text_destination()

    def _choose_pdf_destination(self) -> None:
        """Ask where the PDF should be saved, with safe overwrite handling."""

        self._choose_export_destination(
            title="Export PDF",
            filename=f"{self.path.stem}.pdf",
            filter_name="PDF documents",
            mime_type="application/pdf",
            pattern="*.pdf",
            export_path=_pdf_export_path,
            start_export=self._start_pdf_export,
        )

    def _choose_plain_text_destination(self) -> None:
        """Ask where the Pandoc plain-text export should be saved."""

        self._choose_export_destination(
            title="Export plain text",
            filename=f"{self.path.stem}.txt",
            filter_name="Plain text files",
            mime_type="text/plain",
            pattern="*.txt",
            export_path=_plain_text_export_path,
            start_export=self._start_plain_text_export,
        )

    def _choose_export_destination(
        self,
        *,
        title: str,
        filename: str,
        filter_name: str,
        mime_type: str,
        pattern: str,
        export_path: Callable[[Path], Path],
        start_export: Callable[[Path], None],
    ) -> None:
        """Prompt for an export destination and apply its expected suffix."""

        dialog = Gtk.FileChooserDialog(
            title=title,
            parent=self,
            action=Gtk.FileChooserAction.SAVE,
        )
        dialog.add_buttons(
            "_Cancel", Gtk.ResponseType.CANCEL, "_Export", Gtk.ResponseType.ACCEPT
        )
        dialog.set_do_overwrite_confirmation(True)
        dialog.set_current_folder(str(self.path.parent))
        dialog.set_current_name(filename)
        output_filter = Gtk.FileFilter()
        output_filter.set_name(filter_name)
        output_filter.add_mime_type(mime_type)
        output_filter.add_pattern(pattern)
        dialog.add_filter(output_filter)
        try:
            response = dialog.run()
            filename = dialog.get_filename()
        finally:
            dialog.destroy()
        if response != Gtk.ResponseType.ACCEPT or filename is None:
            self._export_status.set_text("Export cancelled.")
            return
        start_export(export_path(Path(filename)))

    def _start_pdf_export(self, destination: Path) -> None:
        """Run the PDF conversion without blocking the document viewer."""

        if not destination.parent.is_dir():
            self._export_status.set_text("The selected export folder no longer exists.")
            return
        converter = LibreOfficeConverter()
        try:
            paths = converter.prepare(self.path, revision=1)
            process = Gio.Subprocess.new(
                converter.command(paths),
                Gio.SubprocessFlags.STDOUT_PIPE | Gio.SubprocessFlags.STDERR_PIPE,
            )
        except (ConversionError, GLib.Error, OSError) as error:
            converter.close()
            self._export_status.set_text(f"PDF export failed: {error}")
            return

        self._export_converter = converter
        self._export_process = process
        self._export_paths = paths
        self._export_destination = destination
        self._export_status.set_text("Exporting PDF…")
        process.communicate_utf8_async(None, None, self._on_pdf_export_finished, None)

    def _on_pdf_export_finished(
        self,
        process: Gio.Subprocess,
        result: Gio.AsyncResult,
        _data: object,
    ) -> None:
        converter = self._export_converter
        paths = self._export_paths
        destination = self._export_destination
        self._export_converter = None
        self._export_process = None
        self._export_paths = None
        self._export_destination = None
        if converter is None or paths is None or destination is None:
            return

        try:
            _communicated, stdout, stderr = process.communicate_utf8_finish(result)
            pdf = converter.validate(
                paths,
                returncode=process.get_exit_status(),
                stdout=stdout,
                stderr=stderr,
            )
            converter.save_pdf(pdf, destination)
        except (ConversionError, GLib.Error, OSError) as error:
            if not self._closed:
                self._export_status.set_text(f"PDF export failed: {error}")
        else:
            if not self._closed:
                self._export_status.set_text(f"Saved {_compact_path(destination)}")
        finally:
            converter.close()

    def _start_plain_text_export(self, destination: Path) -> None:
        """Run the Pandoc conversion without blocking the document viewer."""

        if not destination.parent.is_dir():
            self._export_status.set_text("The selected export folder no longer exists.")
            return
        converter = PandocConverter()
        try:
            paths = converter.prepare(self.path, revision=1)
            process = Gio.Subprocess.new(
                converter.command(paths),
                Gio.SubprocessFlags.STDOUT_PIPE | Gio.SubprocessFlags.STDERR_PIPE,
            )
        except (ConversionError, GLib.Error, OSError) as error:
            converter.close()
            self._export_status.set_text(f"Plain-text export failed: {error}")
            return

        self._export_converter = converter
        self._export_process = process
        self._export_paths = paths
        self._export_destination = destination
        self._export_status.set_text("Exporting plain text…")
        process.communicate_utf8_async(
            None, None, self._on_plain_text_export_finished, None
        )

    def _on_plain_text_export_finished(
        self,
        process: Gio.Subprocess,
        result: Gio.AsyncResult,
        _data: object,
    ) -> None:
        converter = self._export_converter
        paths = self._export_paths
        destination = self._export_destination
        self._export_converter = None
        self._export_process = None
        self._export_paths = None
        self._export_destination = None
        if (
            not isinstance(converter, PandocConverter)
            or not isinstance(paths, PandocConversionPaths)
            or destination is None
        ):
            return

        try:
            _communicated, stdout, stderr = process.communicate_utf8_finish(result)
            text = converter.validate(
                paths,
                returncode=process.get_exit_status(),
                stdout=stdout,
                stderr=stderr,
            )
            converter.save_text(text, destination)
        except (ConversionError, GLib.Error, OSError) as error:
            if not self._closed:
                self._export_status.set_text(f"Plain-text export failed: {error}")
        else:
            if not self._closed:
                self._export_status.set_text(f"Saved {_compact_path(destination)}")
        finally:
            converter.close()

    def _on_search_changed(self, entry: Gtk.Entry) -> None:
        query = entry.get_text()
        self._search_session_committed = False
        self._clear_search_session_status()
        if not query:
            self._search_matches.clear()
            self._search_index = -1
            self.document.clear_search_highlight()
            self._search_status.set_text("")
            self._queue_search_match_marker_redraw()
            return
        self._search_matches = self.document.search(query)
        if not self._search_matches:
            self._search_index = -1
            self.document.clear_search_highlight()
            self._search_status.set_text("No matches")
            self._queue_search_match_marker_redraw()
            return
        current_page = self.document.current_page_index or 0
        self._search_index = next(
            (
                index
                for index, match in enumerate(self._search_matches)
                if match.page_index >= current_page
            ),
            0,
        )
        self.document.go_to_search_match(self._search_matches[self._search_index])
        self._sync_reading_progress_to_search_match()
        self._update_search_status()

    def _update_search_status(self) -> None:
        self._search_status.set_text(
            f"{self._search_index + 1} of {len(self._search_matches)}"
        )
        if self._search_session_committed:
            self._set_search_session_status(
                f"Search · {self._search_index + 1} of "
                f"{len(self._search_matches)}"
            )
        self._queue_search_match_marker_redraw()

    def _set_search_session_status(self, text: str) -> None:
        """Show the compact search state centred in the bottom status bar."""

        status = getattr(self, "_search_session_status", None)
        if status is None:
            return
        status.set_text(text)
        status.set_no_show_all(False)
        status.show()
        status.set_no_show_all(True)

    def _clear_search_session_status(self) -> None:
        """Remove the committed-search state from the status bar."""

        status = getattr(self, "_search_session_status", None)
        if status is not None:
            status.set_text("")
            status.hide()

    def _search_match_marker_fractions(self) -> list[float]:
        """Return every search result's position in the complete document."""

        if not self._search_matches:
            return []
        return [
            fraction
            for match in self._search_matches
            if (fraction := self.document.search_match_progress_fraction(match))
            is not None
        ]

    def _sync_reading_progress_to_search_match(self) -> None:
        """Settle progress on the same target used by the active match tick."""

        if not 0 <= self._search_index < len(self._search_matches):
            return
        match = self._search_matches[self._search_index]
        fraction = self.document.search_match_progress_fraction(match)
        if fraction is not None:
            self._set_reading_progress_immediately(fraction)

        def settle_after_adjustment() -> bool:
            # Gtk can emit one final scroll notification after the programmatic
            # search jump. Cancel its ordinary reading-progress animation so
            # a centered match and its active tick never diverge.
            if not (
                0 <= self._search_index < len(self._search_matches)
                and self._search_matches[self._search_index] == match
            ):
                return GLib.SOURCE_REMOVE
            settled_fraction = self.document.search_match_progress_fraction(match)
            if settled_fraction is not None:
                self._set_reading_progress_immediately(settled_fraction)
            return GLib.SOURCE_REMOVE

        GLib.idle_add(settle_after_adjustment)

    def _draw_search_match_markers(
        self, marker_layer: Gtk.DrawingArea, context: object
    ) -> bool:
        """Draw a quiet document map over the passive reading-progress rule."""

        fractions = self._search_match_marker_fractions()
        if not fractions:
            return False
        width = marker_layer.get_allocated_width()
        height = marker_layer.get_allocated_height()
        if width <= 0 or height <= 0:
            return False
        palette = _theme_palette(self)
        muted = _rgb(palette["muted"])
        accent = _rgb(palette["accent"])
        for fraction in fractions:
            x = round(fraction * (width - 1))
            context.set_source_rgba(*muted, 0.72)
            context.rectangle(x, 0, 1, height)
            context.fill()
        if 0 <= self._search_index < len(fractions):
            x = round(fractions[self._search_index] * (width - 1))
            context.set_source_rgba(*accent, 1.0)
            context.rectangle(max(0, x - 1), 0, min(2, width), height)
            context.fill()
        return False

    def _queue_search_match_marker_redraw(self) -> None:
        """Refresh the match map while preserving lightweight test doubles."""

        marker_layer = getattr(self, "_search_match_marker_layer", None)
        if marker_layer is not None:
            marker_layer.queue_draw()

    def _move_search_match(self, direction: int) -> bool:
        if not self._search_matches:
            return False
        self._search_index = (self._search_index + direction) % len(
            self._search_matches
        )
        self.document.go_to_search_match(self._search_matches[self._search_index])
        self._sync_reading_progress_to_search_match()
        self._update_search_status()
        return True

    def _on_search_key_press(
        self, _entry: Gtk.Entry, event: Gdk.EventKey
    ) -> bool:
        if event.keyval == Gdk.KEY_Escape:
            self._cancel_search()
            return True
        if event.keyval in {Gdk.KEY_Return, Gdk.KEY_KP_Enter}:
            backwards = bool(event.state & Gdk.ModifierType.SHIFT_MASK)
            if self._move_search_match(-1 if backwards else 1):
                self._search_session_committed = True
                self._update_search_status()
                self._toggle_search()
            return True
        return False

    def _on_page_jump_key_press(self, _entry: Gtk.Entry, event: Gdk.EventKey) -> bool:
        if event.keyval == Gdk.KEY_Escape:
            self._toggle_page_jump()
            return True
        if event.keyval not in {Gdk.KEY_Return, Gdk.KEY_KP_Enter}:
            return False
        try:
            page_number = int(self._page_jump_entry.get_text())
        except ValueError:
            page_number = 0
        if not 1 <= page_number <= self.document.page_count:
            self._page_jump_status.set_text(f"1–{self.document.page_count}")
            return True
        self.document.go_to_page(page_number - 1)
        self._toggle_page_jump()
        return True

    def _url_hint_position(
        self, page: PdfPage, link: UrlLink
    ) -> tuple[int, int] | None:
        """Place a visible PDF URI hint in the unscrolled overlay coordinates."""

        page_x, page_y = page.url_link_position(link)
        translated = page.translate_coordinates(self._url_hint_layer, page_x, page_y)
        if translated is None:
            return None
        if len(translated) == 3:
            success, x, y = translated
            if not success:
                return None
        elif len(translated) == 2:
            x, y = translated
        else:
            return None
        link_width, link_height = page.url_link_size(link)
        layer_width = self._url_hint_layer.get_allocated_width()
        layer_height = self._url_hint_layer.get_allocated_height()
        if (
            layer_width <= 0
            or layer_height <= 0
            or x >= layer_width
            or y >= layer_height
            or x + link_width <= 0
            or y + link_height <= 0
        ):
            return None
        return max(0, x), max(0, y)

    def _set_url_hint_text(self, label: Gtk.Label, code: str) -> None:
        """Render a hint with already typed characters highlighted."""

        label.set_text(" ".join(code.upper()))
        attributes = Pango.AttrList()
        if self._url_hint_prefix:
            color = Gdk.RGBA()
            if color.parse(_theme_palette(self)["highlight"]):
                highlight = Pango.attr_foreground_new(
                    round(color.red * 65535),
                    round(color.green * 65535),
                    round(color.blue * 65535),
                )
                highlight.start_index = 0
                highlight.end_index = len(
                    " ".join(code[: len(self._url_hint_prefix)])
                )
                attributes.insert(highlight)
        label.set_attributes(attributes)

    def _show_url_hints(self) -> None:
        """Label each currently visible URI action with a home-row hint."""

        self._hide_url_hints()
        targets = [
            (link.uri, position)
            for page in self.document._pages
            for link in page._url_links
            if (position := self._url_hint_position(page, link)) is not None
        ]
        for code, (uri, position) in zip(hint_codes(len(targets)), targets, strict=True):
            label = Gtk.Label()
            _style(label, "see-docx-url-hint")
            self._set_url_hint_text(label, code)
            self._url_hint_layer.put(label, *position)
            label.show()
            self._url_hint_targets[code] = uri
            self._url_hint_labels[code] = label
        if self._url_hint_targets:
            self._url_hint_layer.show()

    def _hide_url_hints(self) -> None:
        for label in self._url_hint_labels.values():
            label.destroy()
        self._url_hint_targets.clear()
        self._url_hint_labels.clear()
        self._url_hint_prefix = ""

    def _open_url(self, uri: str) -> None:
        """Launch a selected PDF URI using the desktop's registered handler."""

        try:
            Gio.AppInfo.launch_default_for_uri(uri, None)
        except GLib.Error as error:
            self._set_status(f"Could not open URL: {error.message}")

    def _filter_url_hints(self, character: str) -> None:
        prefix = f"{self._url_hint_prefix}{character}"
        matching = [
            code for code in self._url_hint_targets if code.startswith(prefix)
        ]
        if not matching:
            return
        self._url_hint_prefix = prefix
        for code, label in self._url_hint_labels.items():
            label.set_visible(code in matching)
            if code in matching:
                self._set_url_hint_text(label, code)
        if len(matching) == 1 and matching[0] == prefix:
            uri = self._url_hint_targets[matching[0]]
            self._hide_url_hints()
            self._open_url(uri)

    def _erase_url_hint_character(self) -> None:
        if not self._url_hint_prefix:
            return
        self._url_hint_prefix = self._url_hint_prefix[:-1]
        for code, label in self._url_hint_labels.items():
            visible = code.startswith(self._url_hint_prefix)
            label.set_visible(visible)
            if visible:
                self._set_url_hint_text(label, code)

    def _update_pagination_controls(
        self, page_index: int | None, page_count: int
    ) -> None:
        self._hide_url_hints()
        if page_index is None:
            self._page_indicator.set_text("Page —")
        else:
            self._page_indicator.set_text(f"Page {page_index + 1} / {page_count}")
        adjustment = self.document.widget.get_vadjustment()
        self._animate_reading_progress(
            _reading_progress_fraction(
                adjustment.get_value() - adjustment.get_lower(),
                self.document._maximum_scroll(),
            )
        )
        self._layout_comments()

    def _animate_reading_progress(self, target: float) -> None:
        """Ease the status rule toward the newest exact scroll fraction."""

        target = _clamp_unit(target)
        current = self._reading_progress.get_fraction()
        if abs(target - current) < 0.0001:
            if self._reading_progress_source:
                GLib.source_remove(self._reading_progress_source)
                self._reading_progress_source = 0
            self._reading_progress.set_fraction(target)
            return
        self._reading_progress_start = current
        self._reading_progress_target = target
        self._reading_progress_started_at = GLib.get_monotonic_time()
        if self._reading_progress_source:
            return

        def advance() -> bool:
            elapsed_ms = (
                GLib.get_monotonic_time() - self._reading_progress_started_at
            ) / 1_000
            self._reading_progress.set_fraction(
                _reading_progress_frame(
                    self._reading_progress_start,
                    self._reading_progress_target,
                    elapsed_ms,
                )
            )
            if elapsed_ms >= READING_PROGRESS_DURATION_MS:
                self._reading_progress.set_fraction(self._reading_progress_target)
                self._reading_progress_source = 0
                return GLib.SOURCE_REMOVE
            return GLib.SOURCE_CONTINUE

        self._reading_progress_source = GLib.timeout_add(
            READING_PROGRESS_TICK_MS, advance
        )

    def _set_reading_progress_immediately(self, target: float) -> None:
        """Set search navigation progress without a transient stale frame."""

        if self._reading_progress_source:
            GLib.source_remove(self._reading_progress_source)
            self._reading_progress_source = 0
        self._reading_progress.set_fraction(_clamp_unit(target))

    def _set_status(self, text: str) -> None:
        self._last_status = text

    def _watch_source(self) -> None:
        source_directory = Gio.File.new_for_path(str(self.path.parent))
        try:
            self._monitor = source_directory.monitor_directory(
                Gio.FileMonitorFlags.WATCH_MOVES, None
            )
        except GLib.Error as error:
            self._set_status(f"Live refresh unavailable: {error.message}")
            return
        self._monitor.connect("changed", self._on_source_directory_changed)

    def _is_source(self, file: Gio.File | None) -> bool:
        return bool(file and file.get_basename() == self.path.name)

    def _on_source_directory_changed(
        self,
        _monitor: Gio.FileMonitor,
        file: Gio.File,
        other_file: Gio.File | None,
        event: Gio.FileMonitorEvent,
    ) -> None:
        interesting_events: Iterable[Gio.FileMonitorEvent] = (
            Gio.FileMonitorEvent.CHANGED,
            Gio.FileMonitorEvent.CHANGES_DONE_HINT,
            Gio.FileMonitorEvent.CREATED,
            Gio.FileMonitorEvent.MOVED_IN,
            Gio.FileMonitorEvent.MOVED,
            Gio.FileMonitorEvent.RENAMED,
        )
        if event in interesting_events and (
            self._is_source(file) or self._is_source(other_file)
        ):
            self._queue_refresh()

    def _queue_refresh(self, *, delay: int = REFRESH_DEBOUNCE_MS) -> None:
        if self._closed:
            return
        self._revision += 1
        if self._debounce_source:
            GLib.source_remove(self._debounce_source)
        self._debounce_source = GLib.timeout_add(delay, self._start_refresh)
        if self.document.has_document:
            self._set_status("Source changed; waiting briefly for the new DOCX…")

    def _start_refresh(self) -> bool:
        self._debounce_source = 0
        if self._closed:
            return GLib.SOURCE_REMOVE
        if self._process is not None:
            # The completion callback sees the revision mismatch and runs the
            # latest queued conversion; never show a stale intermediate PDF.
            return GLib.SOURCE_REMOVE
        if not self.path.is_file():
            self._set_status("Waiting for the generated DOCX to reappear…")
            return GLib.SOURCE_REMOVE

        revision = self._revision
        try:
            paths = self._converter.prepare(self.path, revision)
            self._process = Gio.Subprocess.new(
                self._converter.command(paths),
                Gio.SubprocessFlags.STDOUT_PIPE | Gio.SubprocessFlags.STDERR_PIPE,
            )
        except (ConversionError, GLib.Error, OSError) as error:
            self._set_status(f"Preview refresh failed: {error}")
            return GLib.SOURCE_REMOVE

        self._active_revision = revision
        self._active_paths = paths
        self._set_status("Rendering updated DOCX…")
        self._process.communicate_utf8_async(
            None, None, self._on_conversion_finished, None
        )
        return GLib.SOURCE_REMOVE

    def _on_conversion_finished(
        self,
        process: Gio.Subprocess,
        result: Gio.AsyncResult,
        _data: object,
    ) -> None:
        revision = self._active_revision
        paths = self._active_paths
        self._process = None
        self._active_revision = None
        self._active_paths = None
        if revision is None or paths is None:
            return

        try:
            _communicated, stdout, stderr = process.communicate_utf8_finish(result)
            pdf = self._converter.validate(
                paths,
                returncode=process.get_exit_status(),
                stdout=stdout,
                stderr=stderr,
            )
        except (ConversionError, GLib.Error) as error:
            if not self._closed and revision == self._revision:
                self._set_status(f"Preview refresh failed: {error}")
            self._start_latest_if_needed(revision)
            return

        if self._closed:
            self._converter.close()
            return
        if revision != self._revision:
            self._start_latest_if_needed(revision)
            return

        # Capture immediately before replacing existing pages, so scrolling
        # while LibreOffice renders is preserved too. An initial load has no
        # reading location to restore and must not queue a later jump to top.
        had_document = self.document.has_document
        position = self.document.capture_position() if had_document else None
        try:
            document = Poppler.Document.new_from_file(pdf.as_uri(), None)
        except GLib.Error as error:
            self._set_status(f"Preview PDF could not be opened: {error.message}")
            return
        source = _docx_rich_text_source(paths.source_copy)
        self.document.set_document(document, source=source)
        for page in self.document._pages:
            page.connect(
                "size-allocate",
                lambda *_args: GLib.idle_add(self._layout_comments_after_allocate),
            )
        self._set_comments(source.comments if source is not None else ())
        self._update_outline()
        if self._search_panel.get_visible():
            self._on_search_changed(self._search_entry)
        if position is not None:
            GLib.idle_add(self.document.restore_position_after_layout, position)
        else:
            GLib.idle_add(self._fit_document_to_comments)
        self._rendered_revision = revision
        self._converter.discard_before(revision)
        timestamp = datetime.now().strftime("%H:%M:%S")
        self._set_status(f"Live preview updated {timestamp}")

    def _fit_document_to_comments(self) -> bool:
        """Fit the initial page to the width remaining beside the rail."""

        if self.document.has_document and self._comments_visible:
            self.document.fit_to_width()
        return GLib.SOURCE_REMOVE

    def _start_latest_if_needed(self, completed_revision: int) -> None:
        if self._closed:
            return
        if self._revision != completed_revision:
            self._start_refresh()

    def _set_zoom(self, increment: float) -> None:
        self.document.set_zoom(self.document.zoom + increment)

    def _reset_zoom(self) -> None:
        self.document.set_zoom(DEFAULT_ZOOM)

    def _copy_path(self) -> None:
        """Copy the resolved local DOCX path to the regular clipboard."""

        Gtk.Clipboard.get(Gdk.SELECTION_CLIPBOARD).set_text(str(self.path), -1)

    def _copy_all_text(self) -> None:
        """Copy the full document while preserving its original paragraphs."""

        self.document.copy_all_text()

    def _open_in_writer(self) -> None:
        """Hand the current source DOCX to Writer on this Hyprland workspace."""

        if self._writer_handoff_process is not None:
            return
        try:
            process = Gio.Subprocess.new(
                ["hyprctl", "activeworkspace", "-j"],
                Gio.SubprocessFlags.STDOUT_PIPE | Gio.SubprocessFlags.STDERR_PIPE,
            )
        except (GLib.Error, OSError):
            self._launch_writer_direct()
            return
        self._writer_handoff_process = process
        self._set_status("Opening in LibreOffice Writer…")
        process.communicate_utf8_async(
            None,
            None,
            self._on_hyprland_workspace_finished,
            None,
        )

    def _on_hyprland_workspace_finished(
        self,
        process: Gio.Subprocess,
        result: Gio.AsyncResult,
        _data: object,
    ) -> None:
        self._writer_handoff_process = None
        if self._closed:
            return
        try:
            _communicated, stdout, _stderr = process.communicate_utf8_finish(result)
            if process.get_exit_status() != 0:
                self._launch_writer_direct()
                return
            payload = json.loads(stdout)
            workspace = payload.get("id") if isinstance(payload, dict) else None
        except (GLib.Error, OSError, TypeError, ValueError, json.JSONDecodeError):
            self._launch_writer_direct()
            return
        if not isinstance(workspace, int) or isinstance(workspace, bool) or workspace <= 0:
            self._launch_writer_direct()
            return
        self._launch_writer_on_hyprland(workspace)

    def _launch_writer_on_hyprland(self, workspace: int) -> None:
        """Ask Hyprland to launch Writer with a rule for this workspace."""

        command = (
            f"[workspace {workspace} silent] libreoffice --writer "
            f"{shlex.quote(str(self.path))}"
        )
        try:
            process = Gio.Subprocess.new(
                ["hyprctl", "dispatch", "--", "exec", command],
                Gio.SubprocessFlags.STDOUT_PIPE | Gio.SubprocessFlags.STDERR_PIPE,
            )
        except (GLib.Error, OSError):
            self._launch_writer_direct()
            return
        self._writer_handoff_process = process
        process.communicate_utf8_async(
            None,
            None,
            self._on_hyprland_launch_finished,
            None,
        )

    def _on_hyprland_launch_finished(
        self,
        process: Gio.Subprocess,
        result: Gio.AsyncResult,
        _data: object,
    ) -> None:
        self._writer_handoff_process = None
        if self._closed:
            return
        try:
            process.communicate_utf8_finish(result)
            if process.get_exit_status() != 0:
                self._launch_writer_direct()
                return
        except (GLib.Error, OSError):
            self._launch_writer_direct()
            return
        self.close()

    def _launch_writer_direct(self) -> None:
        """Fallback for desktops without a working Hyprland control socket."""

        try:
            Gio.Subprocess.new(
                ["libreoffice", "--writer", str(self.path)],
                Gio.SubprocessFlags.NONE,
            )
        except (GLib.Error, OSError) as error:
            message = error.message if isinstance(error, GLib.Error) else str(error)
            self._set_status(f"Could not open in LibreOffice Writer: {message}")
            return
        self.close()

    def _on_key_press(self, _widget: Gtk.Widget, event: Gdk.EventKey) -> bool:
        # Gtk propagates unhandled entry key events to the window. While the
        # search prompt owns focus, let its normal text editing behavior win
        # over every document-navigation binding.
        if (
            self._search_panel.get_visible()
            and (_widget is self._search_entry or self._search_entry.is_focus())
        ) or (
            self._page_jump_panel.get_visible()
            and (_widget is self._page_jump_entry or self._page_jump_entry.is_focus())
        ):
            return False
        export_panel = getattr(self, "_export_panel", None)
        if export_panel is not None and export_panel.get_visible():
            if event.keyval in {Gdk.KEY_Escape, Gdk.KEY_e}:
                self._toggle_export()
            elif event.keyval == Gdk.KEY_j:
                self._move_export_selection(1)
            elif event.keyval == Gdk.KEY_k:
                self._move_export_selection(-1)
            elif event.keyval in {Gdk.KEY_Return, Gdk.KEY_KP_Enter}:
                self._activate_export_option()
            return True
        if getattr(self, "_url_hint_targets", {}):
            if event.keyval == Gdk.KEY_Escape:
                self._hide_url_hints()
            elif event.keyval == Gdk.KEY_BackSpace:
                self._erase_url_hint_character()
            else:
                character = chr(Gdk.keyval_to_unicode(event.keyval)).lower()
                if character in HINT_CHARS:
                    self._filter_url_hints(character)
            return True
        active_comment_body = getattr(self, "_active_comment_body", None)
        comment_body = active_comment_body() if active_comment_body is not None else None
        if comment_body is not None:
            control = bool(event.state & Gdk.ModifierType.CONTROL_MASK)
            if event.keyval == Gdk.KEY_Escape:
                self._leave_comment_body()
                return True
            if control and event.keyval in {Gdk.KEY_d, Gdk.KEY_u}:
                self._scroll_comment_body(
                    comment_body, 2 if event.keyval == Gdk.KEY_d else -2
                )
                return True
            if not control and event.keyval in {Gdk.KEY_j, Gdk.KEY_k}:
                self._scroll_comment_body(
                    comment_body, 1 if event.keyval == Gdk.KEY_j else -1
                )
                return True
            if not control and event.keyval in {
                Gdk.KEY_Page_Down,
                Gdk.KEY_Page_Up,
            }:
                self._scroll_comment_body(
                    comment_body, 2 if event.keyval == Gdk.KEY_Page_Down else -2
                )
                return True
        if getattr(self, "_comments_focused", False):
            if event.keyval == Gdk.KEY_Escape:
                self._blur_comments()
                return True
            control = bool(event.state & Gdk.ModifierType.CONTROL_MASK)
            if not control:
                if self._comment_body_focused and event.keyval in {
                    Gdk.KEY_g,
                    Gdk.KEY_G,
                }:
                    self._pending_g = False
                    return True
                if not self._comment_body_focused:
                    if not self._comment_annotations:
                        self._pending_g = False
                        return True
                    if self._pending_g:
                        self._pending_g = False
                        if event.keyval == Gdk.KEY_g:
                            self._activate_comment(0, reveal_document=True)
                        return True
                    if event.keyval == Gdk.KEY_g:
                        self._pending_g = True
                        return True
                    if event.keyval == Gdk.KEY_G:
                        self._activate_comment(
                            len(self._comment_annotations) - 1,
                            reveal_document=True,
                        )
                        return True
            if control and event.keyval in {Gdk.KEY_d, Gdk.KEY_u}:
                self._scroll_comment_list(1 if event.keyval == Gdk.KEY_d else -1)
                return True
            if not control and event.keyval in {
                Gdk.KEY_Page_Down,
                Gdk.KEY_Page_Up,
            }:
                self._scroll_comment_list(
                    1 if event.keyval == Gdk.KEY_Page_Down else -1
                )
                return True
            if not control and event.keyval in {Gdk.KEY_j, Gdk.KEY_k}:
                self._move_comment_selection(
                    1 if event.keyval == Gdk.KEY_j else -1
                )
                return True
            if not control and event.keyval in {
                Gdk.KEY_Return,
                Gdk.KEY_KP_Enter,
            }:
                self._enter_comment_body()
                return True
        control = bool(event.state & Gdk.ModifierType.CONTROL_MASK)
        if control:
            self._outline_count = 0
        if control and event.keyval in {Gdk.KEY_d, Gdk.KEY_u}:
            self._pending_g = False
            direction = 1 if event.keyval == Gdk.KEY_d else -1
            if self._outline_panel.get_visible() and (
                _widget is self._outline_tree or self._outline_tree.is_focus()
            ):
                self._move_outline_selection(direction * self._outline_half_page_step())
            else:
                self.document.scroll("half-down" if direction > 0 else "half-up")
            return True
        if control:
            return False
        if self._outline_panel.get_visible():
            if Gdk.KEY_1 <= event.keyval <= Gdk.KEY_9:
                self._outline_count = (
                    self._outline_count * 10 + event.keyval - Gdk.KEY_0
                )
                return True
            if event.keyval == Gdk.KEY_0 and self._outline_count:
                self._outline_count *= 10
                return True
            if event.keyval in {Gdk.KEY_j, Gdk.KEY_k}:
                count = self._outline_count or 1
                self._outline_count = 0
                self._move_outline_selection(
                    count if event.keyval == Gdk.KEY_j else -count
                )
                return True
            self._outline_count = 0
        if self._pending_g:
            self._pending_g = False
            if event.keyval == Gdk.KEY_g:
                self.document.scroll("top")
            return True
        if event.keyval == Gdk.KEY_Tab:
            self._toggle_outline()
        elif event.keyval == Gdk.KEY_c:
            self._focus_comments()
        elif event.keyval == Gdk.KEY_v:
            self._toggle_comments()
        elif event.keyval == Gdk.KEY_W:
            self._open_in_writer()
        elif event.keyval == Gdk.KEY_e:
            self._toggle_export()
        elif event.keyval == Gdk.KEY_slash and not self._search_panel.get_visible():
            self._toggle_search()
        elif event.keyval == Gdk.KEY_colon and not self._page_jump_panel.get_visible():
            self._toggle_page_jump()
        elif (
            event.keyval == Gdk.KEY_n
            and not self._search_panel.get_visible()
            and self._search_matches
        ):
            self._move_search_match(1)
        elif (
            event.keyval == Gdk.KEY_N
            and not self._search_panel.get_visible()
            and self._search_matches
        ):
            self._move_search_match(-1)
        elif self._outline_panel.get_visible() and event.keyval == Gdk.KEY_h:
            self._set_outline_expanded(False)
        elif self._outline_panel.get_visible() and event.keyval == Gdk.KEY_l:
            self._set_outline_expanded(True)
        elif self._outline_panel.get_visible() and event.keyval in {
            Gdk.KEY_Return,
            Gdk.KEY_KP_Enter,
        }:
            self._activate_selected_outline_entry()
        elif event.keyval == Gdk.KEY_g:
            self._pending_g = True
        elif event.keyval == Gdk.KEY_G:
            self.document.scroll("bottom")
        elif event.keyval == Gdk.KEY_Page_Up:
            self.document.go_to_adjacent_page(-1)
        elif event.keyval == Gdk.KEY_Page_Down:
            self.document.go_to_adjacent_page(1)
        elif event.keyval == Gdk.KEY_J:
            self.document.go_to_adjacent_page(1)
        elif event.keyval == Gdk.KEY_K:
            self.document.go_to_adjacent_page(-1)
        elif event.keyval == Gdk.KEY_j:
            self.document.scroll("line-down")
        elif event.keyval == Gdk.KEY_k:
            self.document.scroll("line-up")
        elif event.keyval == Gdk.KEY_a:
            self._copy_all_text()
        elif event.keyval == Gdk.KEY_y:
            self._copy_path()
        elif event.keyval == Gdk.KEY_f:
            self._pending_g = False
            self._show_url_hints()
        elif event.keyval in {Gdk.KEY_plus, Gdk.KEY_KP_Add, Gdk.KEY_equal}:
            self._set_zoom(ZOOM_STEP)
        elif event.keyval in {Gdk.KEY_minus, Gdk.KEY_KP_Subtract}:
            self._set_zoom(-ZOOM_STEP)
        elif event.keyval == Gdk.KEY_0:
            self._reset_zoom()
        elif event.keyval == Gdk.KEY_r:
            self._queue_refresh(delay=0)
        elif event.keyval == Gdk.KEY_q:
            self.close()
        elif event.keyval == Gdk.KEY_Escape:
            self._pending_g = False
            self._cancel_search()
        else:
            return False
        return True

    def _on_delete(self, _window: Gtk.Window, _event: Gdk.EventAny) -> bool:
        self._closed = True
        self._hide_url_hints()
        self.document.cancel_text_selection()
        if self._reading_progress_source:
            GLib.source_remove(self._reading_progress_source)
            self._reading_progress_source = 0
        if self._debounce_source:
            GLib.source_remove(self._debounce_source)
            self._debounce_source = 0
        if self._monitor is not None:
            self._monitor.cancel()
        if self._process is not None:
            self._process.force_exit()
        if self._export_process is not None:
            self._export_process.force_exit()
        if self._writer_handoff_process is not None:
            self._writer_handoff_process.force_exit()
        if self._export_converter is not None:
            self._export_converter.close()
            self._export_converter = None
        # force_exit sends SIGKILL; removing this private temporary profile
        # immediately is safe and prevents abandoned previews after a window is
        # closed mid-conversion.
        self._converter.close()
        return False
