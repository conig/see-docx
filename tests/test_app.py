from __future__ import annotations

from pathlib import Path
import unittest
from unittest.mock import patch

from see_docx.app import main


class AppTests(unittest.TestCase):
    def test_hidden_docx_filename_is_accepted(self) -> None:
        # A DOCX may itself be named ".docx"; pathlib does not treat that
        # complete hidden filename as a suffix, but Writer can open it.
        with patch("see_docx.app.DocxApplication") as application:
            application.return_value.run.return_value = 0

            result = main(["/tmp/.docx"])

        self.assertEqual(result, 0)
        application.assert_called_once_with(Path("/tmp/.docx"))
        application.return_value.run.assert_called_once_with(["see-docx"])


if __name__ == "__main__":
    unittest.main()
