from __future__ import annotations

from pathlib import Path
import tempfile
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
