"""Isolated LibreOffice conversion for the read-only preview pipeline."""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import shutil
import subprocess
import tempfile


class ConversionError(RuntimeError):
    """A DOCX document could not be turned into a preview PDF."""


@dataclass(frozen=True)
class ConversionPaths:
    root: Path
    source_copy: Path
    profile: Path
    output_dir: Path
    pdf: Path


class LibreOfficeConverter:
    """Convert a stable copy of a DOCX file using a private LO profile."""

    def __init__(self) -> None:
        self._root = Path(tempfile.mkdtemp(prefix="see-docx-"))

    @property
    def root(self) -> Path:
        return self._root

    def paths_for(self, revision: int) -> ConversionPaths:
        root = self._root / f"revision-{revision:06d}"
        return ConversionPaths(
            root=root,
            source_copy=root / "source.docx",
            profile=root / "profile",
            output_dir=root / "output",
            pdf=root / "output" / "source.pdf",
        )

    @staticmethod
    def command(paths: ConversionPaths) -> list[str]:
        return [
            "soffice",
            f"-env:UserInstallation={paths.profile.as_uri()}",
            "--headless",
            "--convert-to",
            "pdf:writer_pdf_Export",
            "--outdir",
            str(paths.output_dir),
            str(paths.source_copy),
        ]

    def convert(self, source: Path, revision: int) -> Path:
        """Return a PDF generated from a stable source snapshot.

        A private profile prevents contention with a normal interactive Writer
        session.  Copying first avoids displaying a half-written output while
        Markdown tools replace the generated DOCX atomically.
        """

        paths = self.prepare(source, revision)
        try:
            completed = subprocess.run(
                self.command(paths),
                check=False,
                capture_output=True,
                text=True,
                timeout=60,
            )
        except FileNotFoundError as error:
            raise ConversionError("LibreOffice (soffice) is not installed.") from error
        except subprocess.TimeoutExpired as error:
            raise ConversionError("LibreOffice took more than 60 seconds to render the preview.") from error

        return self.validate(
            paths,
            returncode=completed.returncode,
            stdout=completed.stdout,
            stderr=completed.stderr,
        )

    def prepare(self, source: Path, revision: int) -> ConversionPaths:
        """Copy a stable DOCX snapshot and return its conversion locations."""

        paths = self.paths_for(revision)
        paths.output_dir.mkdir(parents=True)
        before = source.stat()
        shutil.copy2(source, paths.source_copy)
        after = source.stat()
        if (before.st_mtime_ns, before.st_size) != (after.st_mtime_ns, after.st_size):
            raise ConversionError("The source changed while the preview was being prepared.")
        return paths

    @staticmethod
    def validate(
        paths: ConversionPaths,
        *,
        returncode: int,
        stdout: str | None = None,
        stderr: str | None = None,
    ) -> Path:
        if returncode != 0 or not paths.pdf.is_file():
            details = (stderr or stdout or "").strip()
            suffix = f"\n{details}" if details else ""
            raise ConversionError(f"LibreOffice could not render this DOCX.{suffix}")
        return paths.pdf

    def discard_before(self, revision: int) -> None:
        """Keep only the current and later preview work directories."""

        for directory in self._root.glob("revision-*"):
            try:
                number = int(directory.name.removeprefix("revision-"))
            except ValueError:
                continue
            if number < revision:
                shutil.rmtree(directory, ignore_errors=True)

    def close(self) -> None:
        shutil.rmtree(self._root, ignore_errors=True)
