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
| `e`, then `j` / `k` / `Enter` | Open the export tool, choose PDF or plain text, then select its destination |
| Drag across PDF text | Highlight selected glyphs and copy formatted text or table cells immediately to the regular and primary clipboards |
| `f`, then a displayed home-row hint | Open a visible URL with its default desktop application |
| `:number` | Jump directly to one-based page `number` |
| `+` / `-` / `0` | Zoom in / out / reset zoom |
| `r` | Refresh now |
| `q` | Close See DOCX |

The bottom bar shows the resolved document path and current page/total. The
complete document remains scrollable, with every page visible as a separate
sheet above the same background; use the keyboard shortcuts above for page
navigation, zoom, outline navigation, and refresh. The outline comes from
headings exported in the DOCX/PDF; use Word heading styles when a document has
no visible outline.
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
with an accent line. Press `c` to focus the rail without changing the selected
thread. In the thread list, `j` / `k` select the next / previous thread and
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
Pointer text selections continue across page breaks and copy all intersected
pages as one clipboard value. The DOCX source restores formatting and table
structure for Writer-compatible HTML paste, while plain text remains available
for other applications; a selected table cell is copied as a cell rather than
its visual PDF row. Keep holding the mouse button to use the wheel, or drag
above or below a page to auto-scroll while extending the selection.
Press `e` to open the export tool. Choose PDF or plain text with `j` / `k` and
`Enter`, then select a destination in the save dialog. PDF uses an isolated
LibreOffice profile; plain text is exported by Pandoc. Both replace the final
file only after a successful conversion.

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
`pandoc` for plain-text export.

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

Run the unit tests and static compilation check:

```bash
make check
```

The tests cover the position-restoration policy, including unchanged
pagination, pagination reflow, removed pages, and scroll-range clamping. The
conversion command tests confirm each preview uses an isolated LibreOffice
profile and Writer's PDF export filter, and that plain-text export uses Pandoc.

With a desktop display available, run the live-refresh smoke test too:

```bash
pandoc tests/fixtures/live_refresh.md -o /tmp/see-docx-refresh.docx
PYTHONPATH=src python3 tests/ui_refresh_smoke.py /tmp/see-docx-refresh.docx
```

It scrolls the preview, updates a private copy of the supplied DOCX, and
asserts that the page and within-page offset survive the real conversion and
file-monitor refresh.

## Limit of position preservation

The position algorithm is visual rather than semantic: it preserves a page and
its relative offset. If an edit before the current location changes pagination
substantially, the same words may move to another page. Tracking a specific
semantic Markdown source location would require generator-specific source maps
and is intentionally outside this focused viewer.

## License

MIT
