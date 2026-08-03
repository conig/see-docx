from __future__ import annotations

import json
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from see_docx import viewer


STATE_ROLES = {
    "background_alt": "#101112",
    "canvas": "#131415",
    "surface": "#161718",
    "surface_raised": "#191a1b",
    "foreground": "#1c1d1e",
    "foreground_muted": "#1f2021",
    "metadata": "#222324",
    "gtk_command": "#252627",
    "highlight": "#28292a",
    "dim": "#2b2c2d",
    "selection": "#2e2f30",
    "selection_foreground": "#313233",
    "separator": "#343536",
}


class _UnexpectedWidget:
    def get_style_context(self) -> object:
        raise AssertionError("a valid SC1 state must avoid GTK colour lookups")


class _Color:
    def to_string(self) -> str:
        return "#000000"


class _FallbackStyleContext:
    def lookup_color(self, _name: str) -> tuple[bool, _Color]:
        return False, _Color()


class _FallbackWidget:
    def get_style_context(self) -> _FallbackStyleContext:
        return _FallbackStyleContext()


class ThemeStateTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temporary_directory = TemporaryDirectory()
        self.state_path = Path(self.temporary_directory.name) / "current.json"
        self.path_patch = patch.object(
            viewer,
            "_theme_state_path",
            return_value=self.state_path,
        )
        self.path_patch.start()
        viewer._current_theme_state_roles.cache_clear()

    def tearDown(self) -> None:
        viewer._current_theme_state_roles.cache_clear()
        self.path_patch.stop()
        self.temporary_directory.cleanup()

    def _write_state(self, roles: dict[str, str]) -> None:
        self.state_path.write_text(
            json.dumps(
                {
                    "schema_version": 1,
                    "variant": "pink",
                    "roles": roles,
                }
            ),
            encoding="utf-8",
        )

    def test_palette_prefers_the_published_state_file(self) -> None:
        self._write_state(STATE_ROLES)

        palette = viewer._theme_palette(_UnexpectedWidget())

        self.assertEqual(
            palette,
            {
                "background": "#101112",
                "canvas": "#131415",
                "panel": "#191a1b",
                "panel_dark": "#161718",
                "foreground": "#1c1d1e",
                "view_foreground": "#252627",
                "text": "#222324",
                "muted": "#1f2021",
                "metadata": "#222324",
                "accent": "#252627",
                "highlight": "#28292a",
                "accent_dim": "#2b2c2d",
                "selected_background": "#2e2f30",
                "selected_foreground": "#313233",
                "separator": "#343536",
            },
        )

    def test_incomplete_state_falls_back_to_gtk(self) -> None:
        self._write_state({"background_alt": "#101112"})

        palette = viewer._theme_palette(_FallbackWidget())

        self.assertEqual(palette["background"], "#202326")
        self.assertEqual(palette["canvas"], "#30363d")


if __name__ == "__main__":
    unittest.main()
