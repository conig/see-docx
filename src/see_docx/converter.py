"""Isolated LibreOffice conversion for the read-only preview pipeline."""

from __future__ import annotations

from dataclasses import dataclass
import os
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


@dataclass(frozen=True)
class PandocConversionPaths:
    """Private working paths for one Pandoc text-format export."""

    root: Path
    source_copy: Path
    output: Path


def _copy_stable_source(source: Path, destination: Path) -> None:
    """Copy *source* only when it remains unchanged during the copy."""

    before = source.stat()
    shutil.copy2(source, destination)
    after = source.stat()
    if (before.st_mtime_ns, before.st_size) != (after.st_mtime_ns, after.st_size):
        raise ConversionError("The source changed while the export was being prepared.")


def _save_output(source: Path, destination: Path, *, suffix: str) -> Path:
    """Atomically copy a completed export to its requested destination."""

    descriptor, temporary_name = tempfile.mkstemp(
        prefix=f".{destination.stem}-", suffix=suffix, dir=destination.parent
    )
    os.close(descriptor)
    temporary = Path(temporary_name)
    try:
        shutil.copy2(source, temporary)
        temporary.replace(destination)
    except OSError:
        temporary.unlink(missing_ok=True)
        raise
    return destination


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
        _copy_stable_source(source, paths.source_copy)
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

    @staticmethod
    def save_pdf(pdf: Path, destination: Path) -> Path:
        """Atomically publish a rendered PDF at the requested destination.

        The preview conversion itself happens in a private temporary directory.
        Copy through a sibling temporary file so a failed or interrupted export
        never leaves a partially written PDF at the user-selected path.
        """

        return _save_output(pdf, destination, suffix=".pdf")

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


class PandocConverter:
    """Export a stable DOCX snapshot to a Pandoc-supported text format."""

    def __init__(
        self,
        *,
        writer: str = "plain",
        output_suffix: str = ".txt",
        display_name: str = "Plain text",
        writer_options: tuple[str, ...] = (),
    ) -> None:
        self._root = Path(tempfile.mkdtemp(prefix="see-docx-pandoc-"))
        self._writer = writer
        self._output_suffix = output_suffix
        self._display_name = display_name
        self._writer_options = writer_options

    @classmethod
    def markdown(cls) -> PandocConverter:
        """Use Pandoc's Markdown writer with hash-prefixed heading levels."""

        return cls(
            writer="markdown",
            output_suffix=".md",
            display_name="Markdown",
            writer_options=("--markdown-headings=atx",),
        )

    @property
    def display_name(self) -> str:
        return self._display_name

    def paths_for(self, revision: int) -> PandocConversionPaths:
        root = self._root / f"export-{revision:06d}"
        return PandocConversionPaths(
            root=root,
            source_copy=root / "source.docx",
            output=root / f"source{self._output_suffix}",
        )

    def command(self, paths: PandocConversionPaths) -> list[str]:
        return [
            "pandoc",
            "--from=docx",
            f"--to={self._writer}",
            "--wrap=none",
            *self._writer_options,
            "--output",
            str(paths.output),
            str(paths.source_copy),
        ]

    def prepare(self, source: Path, revision: int) -> PandocConversionPaths:
        """Copy a stable DOCX snapshot and return its Pandoc working paths."""

        paths = self.paths_for(revision)
        paths.root.mkdir(parents=True)
        _copy_stable_source(source, paths.source_copy)
        return paths

    def validate(
        self,
        paths: PandocConversionPaths,
        *,
        returncode: int,
        stdout: str | None = None,
        stderr: str | None = None,
    ) -> Path:
        if returncode != 0 or not paths.output.is_file():
            details = (stderr or stdout or "").strip()
            suffix = f"\n{details}" if details else ""
            raise ConversionError(
                f"Pandoc could not export this DOCX as "
                f"{self._display_name.lower()}.{suffix}"
            )
        return paths.output

    def save_output(self, output: Path, destination: Path) -> Path:
        """Atomically publish a completed Pandoc export."""

        return _save_output(output, destination, suffix=self._output_suffix)

    def close(self) -> None:
        shutil.rmtree(self._root, ignore_errors=True)
