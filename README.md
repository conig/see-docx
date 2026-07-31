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
- Displays the PDF pages in a native GTK/Poppler window, floating above a
  contrasting canvas with clear page edges; no editing controls are present.
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
| `Ctrl+d` / `Ctrl+u` | Scroll down / up half a page |
| `Page Down` / `Page Up` | Next / previous document page |
| `gg` / `G` | Jump to the top / bottom |
| `+` / `-` / `0` | Zoom in / out / reset zoom |
| `r` | Refresh now |
| `q` | Close See DOCX |

The toolbar shows the current page and total pages, with previous/next controls
for direct page navigation. The complete document remains scrollable, with
every page visible as a separate sheet above the same background.

## Dependencies

On Arch Linux:

```text
gtk3
python
python-gobject
poppler
libreoffice-fresh
```

The installed application uses `soffice`, which is provided by LibreOffice.

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
conversion command test confirms each preview uses an isolated LibreOffice
profile and Writer's PDF export filter.

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
