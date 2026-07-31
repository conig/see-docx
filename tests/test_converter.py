from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from see_docx.converter import LibreOfficeConverter


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
