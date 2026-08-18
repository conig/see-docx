# See DOCX Specification

This is the behavioral specification for See DOCX. It describes implemented
interaction contracts rather than toolkit-default GTK shortcuts.

## Pointer selection and table clipboard

- A pointer drag through ordinary body text follows document order across
  pages while excluding repeated headers and footers. A selection that begins
  in a running header or footer stays in that flow on its starting page.
- A drag wholly inside one table cell selects only that text range. Its rich
  representation remains one one-cell table rather than losing the source
  cell boundary.
- A double-click inside a table cell selects and visibly highlights the cell's
  complete contents, including when Poppler's interleaved table reading order
  provides no linear source-text matches for that cell or falsely assigns its
  leading rendered line to the cell above. The highlight and clipboard must
  not include any adjacent header or body row.
- While the pointer is over a table, a compact copy control appears wholly
  outside and immediately above-left of that table's top-left corner. It stays
  visible and actionable throughout ordinary pointer travel from any part of
  the table to the control, including whitespace between rendered cell text.
  Clicking it resolves and copies the complete source table through the same
  TSV, single-table HTML, and
  embedded ODF clipboard targets used by table drags. The control is painted
  on the page and must not introduce an input-blocking overlay. For a table
  spanning pages, hovering any rendered page segment exposes the control for
  that segment and the action still copies the one complete source table across
  every page it occupies. Once the clipboard publication succeeds, the
  temporary selection is cleared on every page rather than remaining visibly
  highlighted. A successful control activation
  shows the same compact bottom action notification used by other copy actions,
  confirming that the complete table is ready to paste.
- Once a drag crosses a cell boundary, its endpoint cells define a rectangular
  grid. Horizontal drags select one row; upward and downward vertical drags
  select only their column; diagonal drags select the corresponding rectangle.
  Neither the visible highlight nor the clipboard may include text outside
  that grid. PDF row-major glyph order or an ambiguous repeated-text match must
  not add part of an adjacent cell, including in long documents whose repeated
  rendered headers diverge from body XML.
- Plain table text uses tabs between columns and newlines between rows. Rich
  HTML and embedded ODF each contain exactly one table with the same shape, not
  a sequence of one-cell tables. Pasting a row or column into a blank Writer
  document therefore creates one coherent table.

Receiver behavior is outside the clipboard contract. In LibreOffice Writer
25.8, ordinary paste over an `EntireRow` or `EntireColumn` selection does not
perform one-for-one cell replacement even when the source clipboard is owned
by another Writer process; it appends the complete source selection within
each selected cell. See DOCX guarantees a correctly shaped table object and
TSV fallback, but does not claim to change this Writer behavior.

## Keyboard interaction

### Conventions and dispatch priority

- Letter bindings are case-sensitive: for example, `j` scrolls while `J`
  changes page, and `w` has no binding while `W` opens Writer.
- `Enter` also means keypad Enter. `+` also accepts `=` and keypad `+`; `-`
  also accepts keypad `-`.
- App-defined modes take priority in this order: focused search/page input,
  export chooser, URL hints, focused comment body, focused comment list,
  outline navigation, then document-global bindings.
- Search and page-number inputs retain ordinary GTK text-entry editing keys.
  File chooser shortcuts are likewise supplied by GTK and are outside this
  application-level contract.
- Unless a mode says that it captures all input, global bindings remain
  available when that mode is open.

### Document-global bindings

These bindings apply while reading the document and remain the fallback for
keys not claimed by an active mode.

| Keys | Contract |
| --- | --- |
| `j` / `k` | Scroll down / up by one line step. |
| `Ctrl+d` / `Ctrl+u` | Scroll down / up by half the visible document viewport. |
| `J` / `K` | Go to the next / previous document page without wrapping. |
| `Page Down` / `Page Up` | Go to the next / previous document page without wrapping. |
| `gg` / `G` | Jump to the top / bottom of the complete document. A lone `g` waits for the second key. |
| `Tab` | Toggle the left heading outline. Opening it moves focus into the outline; closing it returns to the document and restores the pre-outline zoom. |
| `c` | Focus the comment rail without changing its selected thread. If comments exist but the rail is hidden, reveal it first. If the document has no comments, do nothing. |
| `v` | Show or hide the right comment rail. If the document has no comments, do nothing. |
| `/` | Open the document-search input. |
| `n` / `N` | Go to the next / previous result of a committed search, wrapping at the ends. With no committed results, leave the key unhandled. |
| `:` | Open the one-based page-number input. |
| `e` | Open the export-format chooser. |
| `f` | Enter URL-hint mode for visible external PDF links. |
| `a` | Copy the complete document with its available rich-text and table structure. |
| `y` | Copy the resolved source DOCX path. |
| `+` / `=` / keypad `+` | Increase manual zoom by 0.10, up to the manual zoom limit. |
| `-` / keypad `-` | Decrease manual zoom by 0.10, down to the manual zoom limit. |
| `Ctrl` + mouse wheel | Increase or decrease zoom by 0.10 at the current reading position; an unmodified wheel keeps normal scrolling behavior. |
| `0` | Reset zoom to the default 1.25 scale. |
| `z` | Zoom in or out so the page fills the width currently allocated to the central document pane, including the effect of either open side panel. Preserve the reading position and do not apply the manual maximum zoom limit. |
| `Z` | Zoom in or out so one complete page fits vertically in the central document pane. Preserve the reading position and do not apply the manual maximum zoom limit. |
| `r` | Queue an immediate source refresh. |
| `W` | Open the source DOCX in LibreOffice Writer and close See DOCX after a successful handoff. On Hyprland, request the current workspace. |
| `Esc` | Cancel the active or committed search session, clear its result state and highlight, and return focus to the document. With no search, clear any pending `g` prefix. |
| `q` | Close See DOCX. |

### Heading-outline mode

`Tab` opens the outline and gives its tree focus. The outline lists visible
heading rows and clamps movement at its first and last entries.

| Keys | Contract |
| --- | --- |
| `j` / `k` | Select the next / previous visible heading row. |
| `[count]j` / `[count]k` | Move down / up by the decimal count; digits `1`–`9` begin a count and `0` extends an existing count. |
| `Ctrl+d` / `Ctrl+u` | With the outline tree focused, move down / up by half the number of currently visible heading rows. |
| `h` / `l` | Collapse / expand the selected heading without changing selection. |
| `Enter` | Jump to the selected heading and briefly mark its exact PDF location. |
| `Tab` | Close the outline and restore the zoom saved when it opened. |

Other unmodified keys fall through to the global bindings. `Ctrl+d` and
`Ctrl+u` scroll the document instead if the outline is visible but its tree
does not own focus.

### Comment-list mode

`c` enters this mode. Selection movement stops at both ends rather than
wrapping and brings an anchored thread into document context.

| Keys | Contract |
| --- | --- |
| `j` / `k` | Select the next / previous comment thread. |
| `gg` / `G` | Select the first / last comment thread. |
| `Ctrl+d` / `Ctrl+u` | Scroll the complete thread list down / up by half its viewport. |
| `Page Down` / `Page Up` | Scroll the complete thread list down / up by half its viewport. |
| `Enter` | Focus the selected thread's conversation body wherever that body is displayed. |
| `Esc` | Leave comment focus and return keyboard focus to the document. |

Global commands not listed above remain available. In particular, `v` hides
the rail and returns focus to the document.

### Comment-body mode

`Enter` from the comment list focuses the active conversation's independently
scrollable body.

| Keys | Contract |
| --- | --- |
| `j` / `k` | Scroll the conversation body down / up by one line step. |
| `Ctrl+d` / `Ctrl+u` | Scroll the conversation body down / up by half its viewport. |
| `Page Down` / `Page Up` | Scroll the conversation body down / up by half its viewport. |
| `Esc` | Return to the comment list while keeping the rail focused. A second `Esc` then returns to the document. |

`g` and `G` are deliberately consumed in the body so they cannot move the
document or change the selected thread. Other unclaimed global commands remain
available.

### Search input and committed search

| Keys | Contract |
| --- | --- |
| `/` | Open the search input and select its current contents. Opening search closes the page-number input. |
| Text input | Search incrementally, select the nearest match at or after the current page, and update the result count. |
| `Enter` | Advance to the next match, commit the search session, close the input, and retain the active highlight and compact result status. If there are no matches, keep the input open. |
| `Shift+Enter` | Move to the previous match and otherwise commit exactly as `Enter` does. |
| `Esc` | Cancel the search completely: clear the query, results, highlight, and committed status, close the input, and focus the document. |
| `n` / `N` | After commit, move to the next / previous match with wraparound. |

While the input owns focus, document-navigation letters are ordinary query
text instead of application commands.

### Page-number input

| Keys | Contract |
| --- | --- |
| `:` | Open an empty one-based page-number input and show the valid range. Opening it closes search input. |
| `Enter` | For a valid integer, place that page at the top of the viewport when possible and close the input. For an invalid or out-of-range value, keep the input open and show the valid range. |
| `Esc` | Close the input without navigating and return focus to the document. |

While this input owns focus, document bindings do not intercept typed text.

### Export chooser

| Keys | Contract |
| --- | --- |
| `e` | Open the chooser at its first format, or close it when already open. Opening it closes search and page-number inputs. |
| `j` / `k` | Select the next / previous export format, wrapping at both ends. |
| `Enter` | Open the destination chooser for the selected format. |
| `Esc` | Close the export chooser and return focus to the document. |

The export chooser captures every key while open; keys not listed above have
no application action.

### URL-hint mode

`f` overlays prefix-free hint codes on visible external links. Codes use the
home-row alphabet `asdfghjkl` and accept either letter case.

| Keys | Contract |
| --- | --- |
| `a`–`l` from `asdfghjkl` | Add the corresponding hint character, filter visible candidates, and open the URI as soon as one complete code matches. |
| `Backspace` | Remove the most recently entered hint character. |
| `Esc` | Dismiss all hints without opening a link. |

URL-hint mode captures every key while active; keys outside its alphabet have
no application action.
