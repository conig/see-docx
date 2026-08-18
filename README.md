# See DOCX

See DOCX is a small, local, read-only DOCX previewer for Markdown-based
workflows. It watches a generated `.docx` file, renders it through LibreOffice
to a private PDF preview, and updates the window when the document changes.

It deliberately is not a Word editor, file picker, office suite, or document
library. It is the document-preview sibling of See Mail: a focused viewer for
the artifact produced by your existing Pandoc, Quarto, R Markdown, or other
build command.

## What it does

- Opens one `.docx` document from the command line or desktop association.
- Renders through LibreOffice, preserving Word page layout more faithfully
  than browser-based DOCX-to-HTML renderers.
- Displays each PDF page as a separate A4 sheet in a vertically scrollable
  native GTK/Poppler print-preview workspace, floating above a contrasting
  canvas with clear page edges; no editing controls are present.
- Watches the output directory and handles atomic file replacement, which is
  how Markdown renderers commonly write a regenerated DOCX.
- Debounces source events for 450 ms and never displays an older conversion
  after a newer source change arrives.
- Keeps the reading position on refresh. Immediately before page replacement,
  it records the current page and fractional vertical position in that page;
  after re-rendering it restores that position. If pagination removes the
  recorded page, it falls back to the equivalent percentage through the whole
  document rather than jumping to the start.
- Converts a private copy using a private LibreOffice profile, so the preview
  does not contend with an interactive Writer session.

The viewer watches the generated DOCX rather than running a project build
itself. Keep your existing Markdown renderer responsible for producing the
DOCX; See DOCX will reflect it as soon as it changes.

## Use

From the checkout:

```bash
PYTHONPATH=src python3 -m see_docx path/to/output.docx
```

After installation:

```bash
see-docx path/to/output.docx
```

## Navigation

The complete mode-by-mode keyboard contract is maintained in
[the specification](SPEC.md#keyboard-interaction). The quick reference below
covers the primary reading workflow.

| Keys | Action |
| --- | --- |
| `j` / `k` | Scroll down / up |
| `J` / `K` | Next / previous document page |
| `Ctrl+d` / `Ctrl+u` | Scroll the document—or move through the open outline—down / up half a page |
| `Ctrl` + mouse wheel | Zoom in / out at the current reading position |
| `Page Down` / `Page Up` | Next / previous document page |
| `gg` / `G` | Jump to the top / bottom |
| `Tab`, then `j` / `k` / `Ctrl+d` / `Ctrl+u` / `h` / `l` / `Enter` | Open the heading outline; select, half-page jump, collapse/expand, or jump to a heading |
| `c` | Focus the comment rail without changing the selected comment |
| `v` | Toggle the right-side DOCX comment rail (shown by default) |
| `W` | Open the source DOCX in LibreOffice Writer and close See DOCX; on Hyprland, keep its current workspace |
| `/`, then `Enter` / `Shift+Enter` / `Esc` | Search document text; commit forward/backward search, or cancel it completely with `Esc` |
| `n` / `N` | Next / previous result while a committed search remains active |
| `a` | Copy the complete document, preserving rich text and tables |
| `y` | Copy the resolved local DOCX path to the clipboard |
| `e`, then `j` / `k` / `Enter` | Open the export tool, choose PDF, plain text, or Markdown, then select its destination |
| Drag across PDF text | Highlight selected glyphs and copy formatted text or table cells immediately to the regular and primary clipboards |
| `f`, then a displayed home-row hint | Open a visible URL with its default desktop application |
| `:number` | Jump directly to one-based page `number` |
| `+` / `-` / `0` | Zoom in / out / reset zoom |
| `z` | Fit the page width to the central pane, accounting for open side panels |
| `Z` | Fit one complete page vertically in the central pane |
| `r` | Refresh now |
| `q` | Close See DOCX |

The bottom bar shows the resolved document path and current page/total. The
complete document remains scrollable, with every page visible as a separate
sheet above the same background; use the keyboard shortcuts above for page
navigation, zoom, outline navigation, and refresh. Resizing the window scales
the current zoom proportionally in either direction, keeping the same balance
between the page width and its surrounding canvas. The outline comes from
headings exported in the DOCX/PDF; use Word heading styles when a document has
no visible outline.
Copying the document path or complete document text shows a compact in-app
confirmation above the bottom bar. Repeated actions replace and reset the same
notification, so confirmations stay visible without accumulating over the
document. The whole-table hover control uses this confirmation as well.
It initially expands only as many whole heading levels as will keep the visible
list below 10 entries; use `l` to reveal deeper structure.
Its relative `j` / `k` offsets use a fixed gutter, like Neovim relative line
numbers. The outline uses one shared spacing unit for its left inset, three-unit
offset gutter, tree indentation, and gaps around expanders, so an offset never
sits between an arrow and its heading.
Expanded or hovered outline arrows use the active SC1GTK variant highlight.
Opening the outline reserves a left column and temporarily fits the complete
page into the remaining viewport; closing it restores the previous zoom.
Selecting a heading places it in reading context and briefly marks its exact
location on the PDF page.
DOCX comments appear in a responsive right-side rail by default. The rail
groups Word comment replies into one anchored thread, showing a root message
followed by indented replies in a single scrollable conversation. Its summary
distinguishes threads from total messages. Each thread shows the author,
messages, and a short quote of the root's attached source text; the quote is
marked on the rendered page and the active thread is connected to its bubble
with an accent line. Selecting a thread with `j` / `k` centres its attached
source text when it has a rendered anchor, including when that text is on
another page. A thread without a source range floats beside the currently
visible page without a misleading connector. When there is enough room, the
original rail card stays in place as an exact-size dashed ghost while an
interactive copy floats beside the page; when there is not enough room,
both the original card and rail layout remain unchanged. Press `c` to focus
the rail without changing the selected thread. In the
thread list, `j` / `k` select the next / previous thread and
stop at the ends; `gg` and `G` jump to the first / last thread. `Ctrl+d` /
`Ctrl+u` scroll the full list by half a viewport. Press `Enter` on the selected
thread to focus its conversation body; `j` / `k` scroll it line by line and
`Ctrl+d` / `Ctrl+u` scroll it by half a viewport. `Esc` returns from the body to
the list, then from the list to the document. The focused rail marks `COMMENTS`
with the active SC1GTK highlight,
and body focus adds a ring to the active thread. Threads share the rail's
right edge; the selected thread widens toward the document instantly, while
leaving the list unfocused clears its active sizing and connector. Press `v`
to slide the rail away when you want a wider reading canvas.
Comment-linked text uses a dimmed variant accent as context and switches to a
high-contrast accent wash and underline only for the focused comment; the
variant highlight role remains reserved for the dark comment rail surfaces.
When the rail has keyboard focus, its selected thread is promoted into a
compact bubble beside the relevant page while its original card becomes the
reserved ghost. Pressing `Enter` focuses and scrolls the conversation wherever
it is currently displayed, without moving a popped-out thread back into the
rail. If the page or anchor is outside the document viewport, or the gutter is
too narrow, no ghost or duplicate is created and the thread stays fully
readable in the rail.
Revealing the rail fits the page width into the remaining document column by
zooming out only when necessary; hiding it restores the previous zoom unless
the user changed zoom manually.
The search prompt appears centered over the document page as one command bar:
`/` marks the mode, the query fills the bar, and `current of total` appears at
the end. After `Enter`, that state becomes a smaller bottom-centre readout,
showing `Search · current of total` until `Esc`, while the page count remains
at the lower right. The
reading-progress rule shows muted ticks for all matches and an accent tick for
the active result; its fill uses the same navigation position as that active
tick. The active text-search result remains highlighted in the PDF while you
move through matches, using the current SC1GTK variant's accent and highlight
colours. Pressing `Esc` cancels that search session: it clears the highlight and
result list, so `n` / `N` no longer navigate until a new search.
Pointer text selections continue across page breaks through the main document
text and copy all intersected body text as one clipboard value, without passing
through repeated headers or footers. A selection that begins in a header or
footer remains confined to that region on its starting page. The DOCX source
restores formatting and table structure for Writer-compatible paste, while
plain text remains available for other applications.

Table selection has an explicit cell-aware contract:

- A drag wholly inside one cell selects only that text range and copies it as
  one one-cell table; it never leaks through neighbouring cells in PDF reading
  order.
- A double-click inside a cell selects only that cell's complete contents,
  without extending into the header or body row above or below it.
- Hovering a table reveals a copy symbol at its top-left. Clicking the symbol
  copies the entire table as one structured table object, confirms success at
  the bottom of the window, and clears the temporary selection.
- Once a drag crosses a cell boundary, the endpoints define a rectangular cell
  grid. A horizontal drag copies one row, and an upward or downward vertical
  drag copies only that column.
- Every such selection is one coherent table object on the rich clipboard, not
  a sequence of unrelated one-cell tables. Pasting into a blank Writer document
  therefore creates one table and retains the selected row/column shape.

LibreOffice Writer 25.8 does not perform one-for-one replacement when an entire
destination row or column is selected: even a row copied natively from one
Writer process is appended in full to every selected destination cell. See DOCX
does not claim to override that receiver-side paste behaviour. Its contract is
that the clipboard contains one correctly shaped table (plus TSV plain text),
rather than the unrelated per-cell tables that caused the original distortion.

Keep holding the mouse button to use the wheel, or drag above or below a page
to auto-scroll while extending the selection.
Press `e` to open the export tool. Choose PDF, plain text, or Markdown with
`j` / `k` and `Enter`, then select a destination in the save dialog. PDF uses
an isolated LibreOffice profile; the text formats are exported by Pandoc.
Markdown uses Pandoc's native Markdown writer with ATX headings, so Word
Heading 1, Heading 2, and later styles become `#`, `##`, and matching
hash-prefixed levels. Every format replaces the final file only after a
successful conversion.

## Dependencies

On Arch Linux:

```text
gtk3
python
python-gobject
poppler
libreoffice-fresh
pandoc
```

The installed application uses `soffice`, which is provided by LibreOffice, and
`pandoc` for plain-text and Markdown export.

## Install

Install for the current user:

```bash
make install
```

This installs the executable to `~/.local/bin`, the Python package to
`~/.local/lib/see-docx`, and a DOCX desktop association to
`~/.local/share/applications`.

Uninstall it with:

```bash
make uninstall
```

## Develop and verify

Run the complete test battery, including the static compilation check, unit
tests, and real GTK pointer/selection/scrolling smoke tests:

```bash
make test-battery
```

`scripts/run-headless-gui-test` creates a disposable 1920x1080 Sway headless
compositor, focuses its workspace 15, then runs any command passed to it. By
default it also blocks access to the user's D-Bus session so portals and
application discovery cannot escape to the live desktop; pass `--session-bus`
when a future test genuinely needs a private session bus. The smoke tests use
Sway's virtual-pointer protocol for real pointer motion, drag, button, and wheel
events, and `wtype` for real keyboard events. On failure the runner uses `grim`
to leave a screenshot, Sway tree, workspace state, config, and log in a
reported directory under `/tmp`.

The graphical kit requires Sway, Grim, wtype, jq, GCC, `wayland-scanner`, the
Wayland client development files, Pandoc, and LibreOffice. It neither maps a
window on the user's compositor nor changes the user's workspace or cursor.
Use `make check` when only the display-independent automated suite is available.
Any additional graphical command can be isolated in the same way:

```bash
scripts/run-headless-gui-test env PYTHONPATH=src python3 tests/my_gui_test.py
```

The tests cover the position-restoration policy, including unchanged
pagination, pagination reflow, removed pages, and scroll-range clamping. The
conversion command tests confirm each preview uses an isolated LibreOffice
profile and Writer's PDF export filter, and that plain-text and Markdown export
use Pandoc's native writers.

With a desktop display available, run the live-refresh smoke test too:

```bash
pandoc tests/fixtures/live_refresh.md -o /tmp/see-docx-refresh.docx
PYTHONPATH=src python3 tests/ui_refresh_smoke.py /tmp/see-docx-refresh.docx
```

It scrolls the preview, updates a private copy of the supplied DOCX, and
asserts that the page and within-page offset survive the real conversion and
file-monitor refresh.

Run the focused-comment smoke test by itself to exercise the realized GTK
overlay and real document input path:

```bash
make comments-smoke
```

It creates a temporary commented DOCX, focuses its comment rail through
`wtype`, and verifies
that the selected thread leaves an exact-size ghost in the rail while a mapped,
interactive copy appears beside the page. With a real compositor pointer it
then checks the PDF text cursor, drag selection, and wheel scrolling before
closing comment focus and confirming that the still-allocated empty layer
passes document input through.

Run the table-selection integration smoke by itself with:

```bash
make rich-selection-smoke
```

It uses real intra-cell, double-click, horizontal, and upward vertical pointer
gestures; checks the plain-text, HTML, and embedded ODF clipboard shapes; then
asks a disposable Writer instance to consume See DOCX's live clipboard in a
blank document as one correctly shaped table. A second long-document fixture
includes repeated running XML before the table so the same gestures cannot
silently fall back to PDF row-major selection. A third open, asymmetric table
forces Poppler's column-major reading order and verifies independently that the
visible final-column highlight cannot enter its wider neighbour or disappear
when a complete geometry-identified cell has no linear source-text matches.
The battery also hovers and clicks the painted whole-table copy control with
real compositor input, including from a later segment of a table spanning
multiple rendered pages, and verifies one complete structured clipboard table
plus cleared selection state on every occupied page after the copy succeeds.

## Limit of position preservation

The position algorithm is visual rather than semantic: it preserves a page and
its relative offset. If an edit before the current location changes pagination
substantially, the same words may move to another page. Tracking a specific
semantic Markdown source location would require generator-specific source maps
and is intentionally outside this focused viewer.

## License

MIT
