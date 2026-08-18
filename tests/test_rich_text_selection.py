from __future__ import annotations

from io import BytesIO
from pathlib import Path
import tempfile
from types import SimpleNamespace
import unittest
import zipfile

import see_docx.viewer as viewer


_DOCUMENT_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:r><w:t>Rich </w:t></w:r>
      <w:r><w:rPr><w:b/></w:rPr><w:t>bold</w:t></w:r>
      <w:r><w:t> and </w:t></w:r>
      <w:r><w:rPr><w:i/></w:rPr><w:t>italic</w:t></w:r>
      <w:r><w:t> prose.</w:t></w:r>
    </w:p>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Product</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Result</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Alpha</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Passed</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
'''

_GRID_DOCUMENT_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B1</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C1</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:rPr><w:b/></w:rPr><w:t>B2</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C2</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>A3</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>B3</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>C3</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
'''

_INTERLEAVED_ROW_DOCUMENT_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:tbl>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Measure status</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>Contribution pathway overview</w:t></w:r></w:p></w:tc>
      </w:tr>
      <w:tr>
        <w:tc><w:p><w:r><w:t>Service focus</w:t></w:r></w:p></w:tc>
        <w:tc><w:p><w:r><w:t>SignalWatch tests a community approach</w:t></w:r></w:p></w:tc>
      </w:tr>
    </w:tbl>
  </w:body>
</w:document>
'''

_RUNNING_MATTER_DOCUMENT_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:document
  xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"
  xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
  <w:body>
    <w:p><w:r><w:t>Body text</w:t></w:r></w:p>
    <w:sectPr>
      <w:headerReference w:type="default" r:id="rIdHeader"/>
      <w:footerReference w:type="default" r:id="rIdFooter"/>
    </w:sectPr>
  </w:body>
</w:document>
'''

_RUNNING_MATTER_RELATIONSHIPS_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship
    Id="rIdHeader"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    Target="header1.xml"/>
  <Relationship
    Id="rIdFooter"
    Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    Target="footer1.xml"/>
</Relationships>
'''

_HEADER_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:hdr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p>
    <w:r><w:t xml:space="preserve">Page </w:t></w:r>
    <w:fldSimple w:instr="PAGE"><w:r><w:t>1</w:t></w:r></w:fldSimple>
  </w:p>
</w:hdr>
'''

_FOOTER_XML = b'''<?xml version="1.0" encoding="UTF-8"?>
<w:ftr xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:p><w:r><w:t>Running footer</w:t></w:r></w:p>
</w:ftr>
'''


class RichTextSelectionTests(unittest.TestCase):
    def _source_path(self, directory: str) -> Path:
        path = Path(directory) / "rich-table.docx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("word/document.xml", _DOCUMENT_XML)
        return path

    def _grid_source(self, directory: str) -> object:
        path = Path(directory) / "grid.docx"
        with zipfile.ZipFile(path, "w") as archive:
            archive.writestr("word/document.xml", _GRID_DOCUMENT_XML)
        source = viewer._docx_rich_text_source(path)
        self.assertIsNotNone(source)
        return source

    def test_selected_docx_runs_are_published_as_rich_html(self) -> None:
        # PDF glyph extraction loses run properties; the source model must put
        # the same bold and italic fragments on the rich clipboard target.
        source_loader = getattr(viewer, "_docx_rich_text_source", None)
        self.assertIsNotNone(source_loader)
        with tempfile.TemporaryDirectory() as directory:
            source = source_loader(self._source_path(directory))

        self.assertIsNotNone(source)
        payload = source.payload_for_text("bold and italic")

        self.assertIsNotNone(payload)
        self.assertEqual(payload.text, "bold and italic")
        self.assertIn("<strong>bold</strong>", payload.html)
        self.assertIn("<em>italic</em>", payload.html)

    def test_selected_table_cell_keeps_its_cell_boundary_and_formatting(self) -> None:
        # Poppler reports one row as `AlphaPassed`; copying the second cell
        # must not turn it into the row or discard its bold run.
        source_loader = getattr(viewer, "_docx_rich_text_source", None)
        self.assertIsNotNone(source_loader)
        with tempfile.TemporaryDirectory() as directory:
            source = source_loader(self._source_path(directory))

        self.assertIsNotNone(source)
        payload = source.payload_for_text("Passed")

        self.assertIsNotNone(payload)
        self.assertEqual(payload.text, "Passed")
        self.assertIn("<table", payload.html)
        self.assertIn("<td", payload.html)
        self.assertIn("<strong>Passed</strong>", payload.html)
        self.assertNotIn("Alpha", payload.html)

    def test_table_drag_snaps_to_complete_cells_and_exports_one_table(self) -> None:
        # Pointer endpoints are glyph positions, but a drag that begins and
        # ends in a table must copy complete cells. A horizontal drag is one
        # tab-delimited row and one HTML table, rather than three unrelated
        # one-cell tables that Writer pastes into every destination cell.
        with tempfile.TemporaryDirectory() as directory:
            source = self._grid_source(directory)

        selection = source.table_selection(
            source._selection_text.index("A2") + 1,
            source._selection_text.index("C2"),
        )

        self.assertIsNotNone(selection)
        self.assertEqual(selection.payload.text, "A2\tB2\tC2")
        self.assertEqual(selection.payload.html.count("<table"), 1)
        self.assertEqual(selection.payload.html.count("<tr>"), 1)
        self.assertEqual(selection.payload.html.count("<td"), 3)
        self.assertIn("<strong>B2</strong>", selection.payload.html)
        self.assertEqual(
            tuple(
                source._selection_text[start:end]
                for start, end in selection.source_ranges
            ),
            ("A2", "B2", "C2"),
        )

    def test_copy_table_action_exports_the_complete_structured_table(self) -> None:
        # The pointer affordance at a table's top-left is one action over the
        # complete source table, not a synthetic drag whose inferred endpoint
        # could omit cells in an irregular or interleaved rendered layout.
        with tempfile.TemporaryDirectory() as directory:
            source = self._grid_source(directory)

        table_index = next(
            fragment.table_index
            for fragment in source._fragments
            if fragment.table_index is not None
        )
        selection = source.table_selection_for_table(table_index)

        self.assertIsNotNone(selection)
        self.assertEqual(
            selection.payload.text,
            "A1\tB1\tC1\nA2\tB2\tC2\nA3\tB3\tC3",
        )
        self.assertEqual(selection.payload.html.count("<table"), 1)
        self.assertEqual(selection.payload.html.count("<tr>"), 3)
        self.assertEqual(selection.payload.html.count("<td"), 9)
        self.assertIsNotNone(selection.payload.odt)
        self.assertEqual(len(selection.source_ranges), 9)

    def test_table_copy_button_sits_wholly_outside_the_table_top_left(
        self,
    ) -> None:
        # The painted control must be reachable without obscuring the table:
        # its hit box stays above-left of the segment with a visible gap and
        # keeps a stable screen size as page zoom changes.
        with tempfile.TemporaryDirectory() as directory:
            source = self._grid_source(directory)
        fragment = next(
            fragment
            for fragment in source._fragments
            if fragment.table_index is not None
        )
        page = SimpleNamespace(
            _zoom=2.0,
            _table_cell_layouts=(
                viewer._TableCellLayout(fragment, 40.0, 30.0, 140.0, 90.0),
            ),
        )

        button = viewer.PdfPage._table_copy_button(
            page, fragment.table_index
        )

        self.assertIsNotNone(button)
        self.assertLessEqual(button.right, 80.0 - 4.0)
        self.assertLessEqual(button.bottom, 60.0 - 4.0)
        self.assertEqual(
            button.right - button.left, viewer.TABLE_COPY_BUTTON_SIZE
        )
        self.assertTrue(
            button.contains(
                (
                    (button.left + button.right) / 2,
                    (button.top + button.bottom) / 2,
                )
            )
        )

    def test_drag_within_one_table_cell_copies_only_its_text_range(self) -> None:
        # An intra-cell drag remains ordinary text selection, but its rich
        # representation is still one coherent one-cell table.
        with tempfile.TemporaryDirectory() as directory:
            source = self._grid_source(directory)

        start = source._selection_text.index("B2")
        fragment = source._table_fragment_at(start)
        selection = source.cell_text_selection(fragment, start + 1, start + 1)

        self.assertIsNotNone(selection)
        self.assertEqual(selection.payload.text, "2")
        self.assertEqual(selection.payload.html.count("<table"), 1)
        self.assertEqual(selection.payload.html.count("<tr>"), 1)
        self.assertEqual(selection.payload.html.count("<td"), 1)
        self.assertIn("<strong>2</strong>", selection.payload.html)

    def test_vertical_table_drag_exports_only_the_selected_column(self) -> None:
        # The source glyph stream is row-major. Selecting B1 down to B3 must
        # not therefore leak C1, A2, C2, or A3 into the copied column.
        with tempfile.TemporaryDirectory() as directory:
            source = self._grid_source(directory)

        selection = source.table_selection(
            source._selection_text.index("B1"),
            source._selection_text.index("B3") + 1,
        )

        self.assertIsNotNone(selection)
        self.assertEqual(selection.payload.text, "B1\nB2\nB3")
        self.assertEqual(selection.payload.html.count("<tr>"), 3)
        self.assertEqual(selection.payload.html.count("<td"), 3)
        self.assertNotIn("A2", selection.payload.html)
        self.assertNotIn("C2", selection.payload.html)
        self.assertEqual(
            tuple(
                source._selection_text[start:end]
                for start, end in selection.source_ranges
            ),
            ("B1", "B2", "B3"),
        )

    def test_table_selection_includes_writer_native_embedded_table(self) -> None:
        # HTML and RTF are generic table imports in Writer: pasting them over
        # an EntireRow selection nests/appends content. The clipboard also
        # needs Writer's embedded ODF table representation for clean cell
        # replacement.
        with tempfile.TemporaryDirectory() as directory:
            source = self._grid_source(directory)

        selection = source.table_selection(
            source._selection_text.index("A2"),
            source._selection_text.index("C2") + 1,
        )

        self.assertIsNotNone(selection)
        self.assertIsNotNone(selection.payload.odt)
        with zipfile.ZipFile(BytesIO(selection.payload.odt)) as archive:
            content = archive.read("content.xml").decode("utf-8")
        self.assertEqual(content.count("<table:table-row>"), 1)
        self.assertEqual(content.count("<table:table-cell"), 3)
        self.assertIn("<table:table-header-rows>", content)
        self.assertIn(">A2</text:p>", content)
        self.assertIn(">B2</text:span>", content)
        self.assertIn('fo:font-weight="bold"', content)
        self.assertIn(">C2</text:p>", content)

    def test_long_rendered_prelude_keeps_late_table_tokens_mapped(self) -> None:
        # Repeated running text is present in the PDF but not in document.xml.
        # Character-level SequenceMatcher autojunk used to lose the late table
        # entirely once this prelude was long enough.
        filler = "ordinary repeated project context continues across the page"
        source_text = "\n".join([filler] * 260) + "\nMeasure of success"
        rendered_text = "\n".join(
            [
                *(
                    f"RUNNING HEADER\n{filler}"
                    for _index in range(260)
                ),
                "Measure of success",
            ]
        )

        mapping = viewer._source_character_matches(rendered_text, source_text)

        self.assertIn(source_text.rindex("Measure"), mapping.values())
        self.assertIn(source_text.rindex("success"), mapping.values())

    def test_table_geometry_recovers_an_unmapped_cell_endpoint(self) -> None:
        # One cell may have no linear source mapping in a complex PDF. Its row
        # and column are still recoverable from mapped neighbouring cells.
        with tempfile.TemporaryDirectory() as directory:
            source = self._grid_source(directory)

        text = source._selection_text
        rectangles = []
        mapping: dict[int, int] = {}
        missing = next(
            fragment
            for fragment in source._fragments
            if fragment.row_index == 2 and fragment.column_start == 1
        )
        for index, character in enumerate(text):
            fragment = source._table_fragment_at(index)
            if fragment is None:
                rectangles.append(SimpleNamespace(x1=0, y1=0, x2=1, y2=1))
                continue
            column = fragment.column_start
            row = fragment.row_index
            rectangles.append(
                SimpleNamespace(
                    x1=column * 100 + index % 2,
                    y1=row * 30,
                    x2=column * 100 + index % 2 + 1,
                    y2=row * 30 + 10,
                )
            )
            if not (missing.start <= index < missing.end) and not character.isspace():
                mapping[index] = index

        layouts = viewer._table_cell_layouts(source, mapping, text, rectangles)
        recovered = next(
            layout.fragment
            for layout in layouts
            if layout.contains((100.5, 65.0))
        )

        self.assertEqual(recovered.row_index, 2)
        self.assertEqual(recovered.column_start, 1)

    def test_row_hit_bands_use_independent_cell_leading_glyphs(self) -> None:
        # LibreOffice can expose the leading line of a tall second-row cell
        # between header glyphs in Poppler's text order.  The linear matcher
        # may then assign that line to the header, but a double-click on the
        # visibly second-row word must still hit only its actual cell.
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "interleaved-row.docx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr(
                    "word/document.xml", _INTERLEAVED_ROW_DOCUMENT_XML
                )
            source = viewer._docx_rich_text_source(path)

        self.assertIsNotNone(source)
        fragments = {
            (fragment.row_index, fragment.column_start): fragment
            for fragment in source._fragments
            if fragment.table_index is not None
        }
        rendered = (
            "Measure status\n"
            "Contribution pathway overview\n"
            "SignalWatch tests a community approach\n"
            "Service focus"
        )
        rectangles = [
            SimpleNamespace(x1=0, y1=0, x2=0, y2=0)
            for _character in rendered
        ]
        mapping: dict[int, int] = {}
        rendered_fragments = (
            ("Measure status", fragments[(0, 0)], 0.0, 0.0),
            ("Contribution pathway overview", fragments[(0, 1)], 100.0, 0.0),
            ("SignalWatch tests a community approach", fragments[(1, 1)], 100.0, 30.0),
            ("Service focus", fragments[(1, 0)], 0.0, 50.0),
        )
        offset = 0
        for text, fragment, left, top in rendered_fragments:
            start = rendered.index(text, offset)
            for character_offset, _character in enumerate(text):
                index = start + character_offset
                glyph_top = (
                    50.0
                    if text.startswith("SignalWatch")
                    and character_offset >= len("SignalWatch tests ")
                    else top
                )
                rectangles[index] = SimpleNamespace(
                    x1=left + character_offset,
                    y1=glyph_top,
                    x2=left + character_offset + 1,
                    y2=glyph_top + 10,
                )
                if text.startswith("SignalWatch") and character_offset < len(
                    "SignalWatch tests "
                ):
                    # This is the observed bad alignment: the target cell's
                    # first rendered line extends the inferred header row.
                    header = fragments[(0, 1)]
                    mapping[index] = header.start + min(
                        character_offset, header.end - header.start - 1
                    )
                else:
                    mapping[index] = fragment.start + character_offset
            offset = start + len(text)

        layouts = viewer._table_cell_layouts(
            source, mapping, rendered, rectangles
        )
        recovered = next(
            layout.fragment
            for layout in layouts
            if layout.contains((104.0, 35.0))
        )

        self.assertEqual(recovered.row_index, 1)
        self.assertEqual(recovered.column_start, 1)

    def test_column_hit_bands_follow_the_next_cells_leading_edge(self) -> None:
        # Wide borderless cells may leave a large blank gap after one cell's
        # last mapped word. The midpoint of that gap lies inside the first
        # cell, so the following column begins at its own leading edge minus
        # cell padding instead.
        bands = viewer._table_column_bands(
            {
                (0, 0, 1): (10.0, 30.0),
                (0, 1, 2): (40.0, 70.0),
                (0, 2, 3): (120.0, 160.0),
            }
        )

        self.assertEqual(bands[(0, 0, 1)], (4.0, 34.0))
        self.assertEqual(bands[(0, 1, 2)], (34.0, 114.0))
        self.assertEqual(bands[(0, 2, 3)], (114.0, 166.0))

    def test_complete_cell_highlight_prefers_geometry_over_a_false_text_match(
        self,
    ) -> None:
        # Repeated table prose can align a glyph in one column to an identical
        # source token in the next. Once a complete cell has been selected,
        # the inferred two-dimensional cell geometry is authoritative: the
        # false linear match must not paint part of the adjacent column.
        with tempfile.TemporaryDirectory() as directory:
            source = self._grid_source(directory)

        first = source._table_fragment_at(source._selection_text.index("A1"))
        final = source._table_fragment_at(source._selection_text.index("C1"))
        self.assertIsNotNone(first)
        self.assertIsNotNone(final)
        rectangles = [
            SimpleNamespace(x1=0, y1=0, x2=10, y2=10),
            SimpleNamespace(x1=20, y1=0, x2=30, y2=10),
        ]
        rendered_page = SimpleNamespace(
            get_text=lambda: "AC",
            get_text_layout=lambda: (True, rectangles),
        )
        page = SimpleNamespace(
            _page=rendered_page,
            _text_selection_start=None,
            _text_selection_end=None,
            _text_selection_flow="main",
            _selection_flow_map={},
            _text_selection_source_ranges=((final.start, final.end),),
            # Both glyphs appear to map into C1, reproducing a false repeated-
            # token match for the glyph that is visibly inside A1.
            _source_character_map={0: final.start, 1: final.start + 1},
            _table_fragment_map={0: first, 1: final},
        )

        selected_text, selected_rectangles = viewer.PdfPage._layout_text_selection(
            page,
            SimpleNamespace(left=0, top=0, right=30, bottom=10),
        )

        self.assertEqual(selected_text, "C")
        self.assertEqual(selected_rectangles, [rectangles[1]])

    def test_running_header_and_footer_text_are_separate_selection_flows(
        self,
    ) -> None:
        # Running matter renders into the PDF but is not part of document.xml's
        # body flow, so retain it separately for semantic pointer selection.
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / "running-matter.docx"
            with zipfile.ZipFile(path, "w") as archive:
                archive.writestr("word/document.xml", _RUNNING_MATTER_DOCUMENT_XML)
                archive.writestr(
                    "word/_rels/document.xml.rels",
                    _RUNNING_MATTER_RELATIONSHIPS_XML,
                )
                archive.writestr("word/header1.xml", _HEADER_XML)
                archive.writestr("word/footer1.xml", _FOOTER_XML)

            source = viewer._docx_rich_text_source(path)

        self.assertIsNotNone(source)
        self.assertEqual(
            source._selection_flow_texts,
            {
                "header": ("Page 1", "Page "),
                "footer": ("Running footer",),
            },
        )
