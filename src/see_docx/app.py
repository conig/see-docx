"""Application entry point for See DOCX."""

from __future__ import annotations

import argparse
from pathlib import Path
import sys

import gi

gi.require_version("Gtk", "3.0")

from gi.repository import Gio, Gtk

from . import __version__
from .viewer import APPLICATION_ID, DocxWindow


class DocxApplication(Gtk.Application):
    def __init__(self, path: Path) -> None:
        super().__init__(
            application_id=APPLICATION_ID,
            flags=Gio.ApplicationFlags.NON_UNIQUE,
        )
        self.path = path
        self.window: DocxWindow | None = None

    def do_activate(self) -> None:
        if self.window is None:
            self.window = DocxWindow(self, self.path)
        self.window.show_all()
        self.window.present()


def _parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="see-docx",
        description="View one DOCX file and refresh when it is regenerated.",
    )
    parser.add_argument("path", type=Path, help="DOCX file to preview")
    parser.add_argument("--version", action="version", version=f"see-docx {__version__}")
    return parser


def main(argv: list[str] | None = None) -> int:
    arguments = _parser().parse_args(argv)
    path = arguments.path.expanduser()
    if path.suffix.lower() != ".docx":
        print("see-docx expects a .docx file.", file=sys.stderr)
        return 2
    application = DocxApplication(path)
    return application.run(["see-docx"])
