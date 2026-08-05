from __future__ import annotations

from pathlib import Path
import unittest

from see_docx.viewer import _compact_path


class PathTests(unittest.TestCase):
    def test_compacts_only_the_current_users_home_prefix(self) -> None:
        home = Path("/home/marine")

        self.assertEqual(_compact_path(home / "reports" / "output.docx", home), "~/reports/output.docx")
        self.assertEqual(_compact_path(home, home), "~")
        self.assertEqual(_compact_path(Path("/tmp/output.docx"), home), "/tmp/output.docx")
