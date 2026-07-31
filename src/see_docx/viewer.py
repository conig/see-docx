"""GTK/Poppler document window with live DOCX refresh."""

from __future__ import annotations

from collections.abc import Callable, Iterable
from datetime import datetime
from pathlib import Path

import gi

gi.require_version("Gdk", "3.0")
gi.require_version("Gtk", "3.0")
gi.require_version("Poppler", "0.18")

from gi.repository import Gdk, Gio, GLib, Gtk, Pango, Poppler

from .converter import ConversionError, ConversionPaths, LibreOfficeConverter
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
MAX_ZOOM = 2.00
ZOOM_STEP = 0.10
REFRESH_DEBOUNCE_MS = 450
SCROLL_STEP = 56


def _style(widget: Gtk.Widget, class_name: str) -> None:
    widget.get_style_context().add_class(class_name)


def _label(text: str, *, xalign: float = 0.0) -> Gtk.Label:
    return Gtk.Label(label=text, xalign=xalign)


def _lookup_color(widget: Gtk.Widget, name: str, fallback: str) -> str:
    found, color = widget.get_style_context().lookup_color(name)
    return color.to_string() if found else fallback


def _theme_palette(widget: Gtk.Widget) -> dict[str, str]:
    """Resolve SC1 semantic colours, with useful standard GTK fallbacks."""

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
        "text": _lookup_color(widget, "theme_text_color", foreground),
        "muted": _lookup_color(widget, "sc1-fg-muted", "#a9b2bd"),
        "accent": accent,
        "highlight": _lookup_color(widget, "sc1-command-highlight", foreground),
        "accent_dim": _lookup_color(widget, "sc1-command-dim", accent),
        "selected_background": _lookup_color(
            widget, "sc1-selection-bg-solid", selected_background
        ),
        "separator": _lookup_color(
            widget, "sc1-separator", _lookup_color(widget, "borders", "#3b4148")
        ),
    }


def _app_css(widget: Gtk.Widget) -> bytes:
    palette = _theme_palette(widget)
    return f"""
window.see-docx-window {{
  background-color: {palette["background"]};
  color: {palette["foreground"]};
}}
.see-docx-root {{ background-color: {palette["background"]}; }}
.see-docx-header {{
  background-color: {palette["panel"]};
  border-left: 4px solid {palette["accent"]};
  border-bottom: 1px solid {palette["separator"]};
  padding: 13px 18px 14px;
}}
.see-docx-eyebrow {{
  color: {palette["accent"]};
  font-size: 0.80em;
  font-weight: 700;
  letter-spacing: 0.08em;
}}
.see-docx-filename {{
  color: {palette["highlight"]};
  font-size: 16pt;
  font-weight: 700;
  margin-top: 2px;
}}
.see-docx-path {{
  color: {palette["muted"]};
  font-size: 0.87em;
  margin-top: 3px;
}}
.see-docx-toolbar button {{
  background-color: {palette["panel_dark"]};
  border-color: {palette["separator"]};
  color: {palette["text"]};
  min-height: 24px;
  padding: 2px 10px;
}}
.see-docx-toolbar button:hover {{
  background-color: {palette["selected_background"]};
  border-color: {palette["accent_dim"]};
  color: {palette["highlight"]};
}}
.see-docx-page-indicator {{
  color: {palette["muted"]};
  font-size: 0.87em;
}}
.see-docx-workspace,
.see-docx-workspace viewport {{ background-color: {palette["canvas"]}; }}
.see-docx-pages {{ background-color: {palette["canvas"]}; padding: 32px 40px; }}
.see-docx-page {{
  background-color: #ffffff;
  border: 1px solid {palette["separator"]};
  box-shadow: 0 5px 20px rgba(0, 0, 0, 0.32);
  margin-bottom: 18px;
}}
.see-docx-status {{
  background-color: {palette["panel_dark"]};
  border-top: 1px solid {palette["separator"]};
  color: {palette["muted"]};
  font-size: 0.87em;
  padding: 8px 12px;
}}
""".encode()


class PdfPage(Gtk.DrawingArea):
    """A lazily drawn Poppler page at the current zoom level."""

    def __init__(self, page: Poppler.Page, zoom: float) -> None:
        super().__init__()
        self._page = page
        self._zoom = zoom
        self._width, self._height = page.get_size()
        self.set_app_paintable(True)
        self.set_halign(Gtk.Align.CENTER)
        _style(self, "see-docx-page")
        self._resize()
        self.connect("draw", self._on_draw)

    def _resize(self) -> None:
        self.set_size_request(
            max(1, round(self._width * self._zoom)),
            max(1, round(self._height * self._zoom)),
        )

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
        return False


class PdfDocumentView:
    """A scrollable PDF page stack that can restore a document location."""

    def __init__(
        self,
        on_page_changed: Callable[[int | None, int], None] | None = None,
    ) -> None:
        self.widget = Gtk.ScrolledWindow()
        self.widget.set_policy(Gtk.PolicyType.AUTOMATIC, Gtk.PolicyType.AUTOMATIC)
        self.widget.set_kinetic_scrolling(True)
        _style(self.widget, "see-docx-workspace")

        self._pages_box = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        self._pages_box.set_halign(Gtk.Align.CENTER)
        self._pages_box.set_valign(Gtk.Align.START)
        _style(self._pages_box, "see-docx-pages")
        self.widget.add(self._pages_box)
        self._pages: list[PdfPage] = []
        self._document: Poppler.Document | None = None
        self._pending_restore: DocumentPosition | None = None
        self._restore_attempts = 0
        self._on_page_changed = on_page_changed
        self.zoom = DEFAULT_ZOOM
        self.widget.get_vadjustment().connect("value-changed", self._on_scroll_changed)

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

    def _on_scroll_changed(self, _adjustment: Gtk.Adjustment) -> None:
        self._notify_page_changed()

    def _notify_page_changed(self) -> None:
        if self._on_page_changed is not None:
            self._on_page_changed(self.current_page_index, self.page_count)

    def set_document(self, document: Poppler.Document) -> None:
        for page in self._pages:
            self._pages_box.remove(page)
        self._pages.clear()
        self._document = document
        for number in range(document.get_n_pages()):
            page = PdfPage(document.get_page(number), self.zoom)
            self._pages_box.pack_start(page, False, False, 0)
            self._pages.append(page)
        self._pages_box.show_all()
        self._notify_page_changed()

    def set_zoom(self, zoom: float) -> None:
        if self._document is None:
            return
        zoom = min(max(zoom, MIN_ZOOM), MAX_ZOOM)
        if abs(zoom - self.zoom) < 0.001:
            return
        position = self.capture_position()
        self.zoom = zoom
        self.set_document(self._document)
        GLib.idle_add(self.restore_position_after_layout, position)

    def _page_geometries(self) -> list[PageGeometry]:
        geometries: list[PageGeometry] = []
        for page in self._pages:
            allocation = page.get_allocation()
            geometries.append(PageGeometry(float(allocation.y), float(allocation.height)))
        return geometries

    def _maximum_scroll(self) -> float:
        adjustment = self.widget.get_vadjustment()
        return max(0.0, adjustment.get_upper() - adjustment.get_lower() - adjustment.get_page_size())

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


class DocxWindow(Gtk.ApplicationWindow):
    """A read-only, self-refreshing DOCX preview window."""

    def __init__(self, application: Gtk.Application, path: Path) -> None:
        super().__init__(application=application)
        self.path = path.expanduser().resolve()
        self._converter = LibreOfficeConverter()
        self._monitor: Gio.FileMonitor | None = None
        self._process: Gio.Subprocess | None = None
        self._active_paths: ConversionPaths | None = None
        self._active_revision: int | None = None
        self._revision = 0
        self._rendered_revision = 0
        self._debounce_source = 0
        self._closed = False
        self._pending_g = False
        self._previous_page_button: Gtk.Button
        self._next_page_button: Gtk.Button
        self._page_indicator: Gtk.Label

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

        header = Gtk.HeaderBar()
        header.set_show_close_button(True)
        header.set_title("See DOCX")
        header.set_subtitle(self.path.name)
        self.set_titlebar(header)

        root = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        _style(root, "see-docx-root")
        root.pack_start(self._build_document_header(), False, False, 0)
        self.document = PdfDocumentView(self._update_pagination_controls)
        root.pack_start(self.document.widget, True, True, 0)
        self.status = _label("Preparing live preview…")
        _style(self.status, "see-docx-status")
        root.pack_end(self.status, False, False, 0)
        self.add(root)

        self._watch_source()
        self._queue_refresh(delay=0)

    def _build_document_header(self) -> Gtk.Widget:
        container = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=16)
        _style(container, "see-docx-header")

        details = Gtk.Box(orientation=Gtk.Orientation.VERTICAL)
        eyebrow = _label("READ-ONLY LIVE PREVIEW")
        _style(eyebrow, "see-docx-eyebrow")
        details.pack_start(eyebrow, False, False, 0)
        filename = _label(self.path.name)
        filename.set_ellipsize(Pango.EllipsizeMode.MIDDLE)
        _style(filename, "see-docx-filename")
        details.pack_start(filename, False, False, 0)
        source = _label(str(self.path))
        source.set_ellipsize(Pango.EllipsizeMode.MIDDLE)
        _style(source, "see-docx-path")
        details.pack_start(source, False, False, 0)
        container.pack_start(details, True, True, 0)

        toolbar = Gtk.Box(orientation=Gtk.Orientation.HORIZONTAL, spacing=5)
        _style(toolbar, "see-docx-toolbar")
        self._previous_page_button = Gtk.Button(label="‹")
        self._previous_page_button.set_tooltip_text("Previous page (Page Up)")
        self._previous_page_button.set_sensitive(False)
        self._previous_page_button.connect(
            "clicked", lambda _button: self.document.go_to_adjacent_page(-1)
        )
        toolbar.pack_start(self._previous_page_button, False, False, 0)

        self._page_indicator = _label("Page —", xalign=0.5)
        self._page_indicator.set_width_chars(12)
        _style(self._page_indicator, "see-docx-page-indicator")
        toolbar.pack_start(self._page_indicator, False, False, 0)

        self._next_page_button = Gtk.Button(label="›")
        self._next_page_button.set_tooltip_text("Next page (Page Down)")
        self._next_page_button.set_sensitive(False)
        self._next_page_button.connect(
            "clicked", lambda _button: self.document.go_to_adjacent_page(1)
        )
        toolbar.pack_start(self._next_page_button, False, False, 0)

        for label, callback, tooltip in (
            ("−", lambda _button: self._set_zoom(-ZOOM_STEP), "Zoom out (-)"),
            ("100%", lambda _button: self._reset_zoom(), "Reset zoom (0)"),
            ("+", lambda _button: self._set_zoom(ZOOM_STEP), "Zoom in (+)"),
            ("Refresh", lambda _button: self._queue_refresh(delay=0), "Refresh now (r)"),
        ):
            button = Gtk.Button(label=label)
            button.set_tooltip_text(tooltip)
            button.connect("clicked", callback)
            toolbar.pack_start(button, False, False, 0)
        container.pack_end(toolbar, False, False, 0)
        return container

    def _update_pagination_controls(
        self, page_index: int | None, page_count: int
    ) -> None:
        if page_index is None:
            self._page_indicator.set_text("Page —")
        else:
            self._page_indicator.set_text(f"Page {page_index + 1} of {page_count}")
        self._previous_page_button.set_sensitive(page_index is not None and page_index > 0)
        self._next_page_button.set_sensitive(
            page_index is not None and page_index < page_count - 1
        )

    def _set_status(self, text: str) -> None:
        self.status.set_text(text)

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
        if event in interesting_events and (self._is_source(file) or self._is_source(other_file)):
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
        self._process.communicate_utf8_async(None, None, self._on_conversion_finished, None)
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
        self.document.set_document(document)
        if position is not None:
            GLib.idle_add(self.document.restore_position_after_layout, position)
        self._rendered_revision = revision
        self._converter.discard_before(revision)
        timestamp = datetime.now().strftime("%H:%M:%S")
        page_word = "page" if self.document.page_count == 1 else "pages"
        self._set_status(
            f"Live preview updated {timestamp} · {self.document.page_count} {page_word} · "
            "position preserved"
        )

    def _start_latest_if_needed(self, completed_revision: int) -> None:
        if self._closed:
            return
        if self._revision != completed_revision:
            self._start_refresh()

    def _set_zoom(self, increment: float) -> None:
        self.document.set_zoom(self.document.zoom + increment)

    def _reset_zoom(self) -> None:
        self.document.set_zoom(DEFAULT_ZOOM)

    def _on_key_press(self, _widget: Gtk.Widget, event: Gdk.EventKey) -> bool:
        control = bool(event.state & Gdk.ModifierType.CONTROL_MASK)
        if control and event.keyval in {Gdk.KEY_d, Gdk.KEY_u}:
            self._pending_g = False
            self.document.scroll("half-down" if event.keyval == Gdk.KEY_d else "half-up")
            return True
        if control:
            return False
        if self._pending_g:
            self._pending_g = False
            if event.keyval == Gdk.KEY_g:
                self.document.scroll("top")
            return True
        if event.keyval == Gdk.KEY_g:
            self._pending_g = True
        elif event.keyval == Gdk.KEY_G:
            self.document.scroll("bottom")
        elif event.keyval == Gdk.KEY_Page_Up:
            self.document.go_to_adjacent_page(-1)
        elif event.keyval == Gdk.KEY_Page_Down:
            self.document.go_to_adjacent_page(1)
        elif event.keyval == Gdk.KEY_j:
            self.document.scroll("line-down")
        elif event.keyval == Gdk.KEY_k:
            self.document.scroll("line-up")
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
        else:
            return False
        return True

    def _on_delete(self, _window: Gtk.Window, _event: Gdk.EventAny) -> bool:
        self._closed = True
        if self._debounce_source:
            GLib.source_remove(self._debounce_source)
            self._debounce_source = 0
        if self._monitor is not None:
            self._monitor.cancel()
        if self._process is not None:
            self._process.force_exit()
        # force_exit sends SIGKILL; removing this private temporary profile
        # immediately is safe and prevents abandoned previews after a window is
        # closed mid-conversion.
        self._converter.close()
        return False
