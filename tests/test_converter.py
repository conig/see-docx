from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from see_docx.converter import LibreOfficeConverter, PandocConverter


class ConverterTests(unittest.TestCase):
    def test_uses_a_private_profile_and_pdf_export_filter(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            converter = LibreOfficeConverter()
            try:
                paths = converter.paths_for(7)
                command = converter.command(paths)
            finally:
                converter.close()

        self.assertEqual(command[0], "soffice")
        self.assertIn("--headless", command)
        self.assertIn("pdf:writer_pdf_Export", command)
        self.assertIn("--outdir", command)
        self.assertTrue(command[1].startswith("-env:UserInstallation=file://"))
        self.assertTrue(command[-1].endswith("revision-000007/source.docx"))

    def test_saves_a_pdf_by_atomically_replacing_the_destination(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            rendered_pdf = root / "rendered.pdf"
            destination = root / "report.pdf"
            rendered_pdf.write_bytes(b"new PDF")
            destination.write_bytes(b"old PDF")

            saved = LibreOfficeConverter.save_pdf(rendered_pdf, destination)

            self.assertEqual(saved, destination)
            self.assertEqual(destination.read_bytes(), b"new PDF")
            self.assertEqual(list(root.glob(".report-*.pdf")), [])

    def test_uses_pandoc_to_export_plain_text_from_a_stable_docx_copy(self) -> None:
        converter = PandocConverter()
        try:
            paths = converter.paths_for(7)
            command = converter.command(paths)
        finally:
            converter.close()

        self.assertEqual(command[0], "pandoc")
        self.assertIn("--from=docx", command)
        self.assertIn("--to=plain", command)
        self.assertEqual(command[command.index("--output") + 1], str(paths.text))
        self.assertTrue(command[-1].endswith("export-000007/source.docx"))

    def test_saves_plain_text_by_atomically_replacing_the_destination(self) -> None:
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory)
            rendered_text = root / "rendered.txt"
            destination = root / "report.txt"
            rendered_text.write_text("new text\n", encoding="utf-8")
            destination.write_text("old text\n", encoding="utf-8")

            saved = PandocConverter.save_text(rendered_text, destination)

            self.assertEqual(saved, destination)
            self.assertEqual(destination.read_text(encoding="utf-8"), "new text\n")
            self.assertEqual(list(root.glob(".report-*.txt")), [])
