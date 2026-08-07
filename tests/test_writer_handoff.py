from __future__ import annotations

from pathlib import Path
from types import SimpleNamespace
import unittest
from unittest.mock import call, patch

from see_docx.viewer import DocxWindow, Gdk, Gio


class _HiddenPanel:
    def get_visible(self) -> bool:
        return False


class _WriterKeyWindow:
    _on_key_press = DocxWindow._on_key_press

    def __init__(self) -> None:
        self._search_panel = _HiddenPanel()
        self._page_jump_panel = _HiddenPanel()
        self._export_panel = _HiddenPanel()
        self._outline_panel = _HiddenPanel()
        self._search_entry = SimpleNamespace(is_focus=lambda: False)
        self._page_jump_entry = SimpleNamespace(is_focus=lambda: False)
        self._url_hint_targets: dict[str, str] = {}
        self._pending_g = False
        self.opened = 0

    def _open_in_writer(self) -> None:
        self.opened += 1


class _FakeProcess:
    def __init__(self, stdout: str = "", status: int = 0) -> None:
        self.stdout = stdout
        self.status = status

    def communicate_utf8_async(self, _stdin, _cancellable, callback, data) -> None:
        callback(self, SimpleNamespace(), data)

    def communicate_utf8_finish(self, _result):
        return True, self.stdout, ""

    def get_exit_status(self) -> int:
        return self.status


class _WriterWindow:
    def __init__(self, statuses: list[str], closed: list[bool]) -> None:
        self.path = Path("/tmp/review.docx")
        self._writer_handoff_process = None
        self._closed = False
        self._statuses = statuses
        self._closed = closed

    def close(self) -> None:
        self._closed.append(True)

    def _set_status(self, text: str) -> None:
        self._statuses.append(text)

    def _launch_writer_direct(self) -> None:
        DocxWindow._launch_writer_direct(self)

    def _launch_writer_on_hyprland(self, workspace: int) -> None:
        DocxWindow._launch_writer_on_hyprland(self, workspace)

    def _on_hyprland_workspace_finished(self, process, result, data) -> None:
        DocxWindow._on_hyprland_workspace_finished(self, process, result, data)

    def _on_hyprland_launch_finished(self, process, result, data) -> None:
        DocxWindow._on_hyprland_launch_finished(self, process, result, data)


class WriterHandoffTests(unittest.TestCase):
    def test_uppercase_W_is_the_writer_handoff_key(self) -> None:
        window = _WriterKeyWindow()

        self.assertTrue(
            window._on_key_press(
                window, SimpleNamespace(keyval=Gdk.KEY_W, state=0)
            )
        )
        self.assertEqual(window.opened, 1)

    def test_opens_the_source_docx_in_writer_then_closes_preview(self) -> None:
        closed: list[bool] = []
        window = _WriterWindow([], closed)
        workspace_process = _FakeProcess('{"id": 7}')
        launch_process = _FakeProcess("ok")

        with patch(
            "see_docx.viewer.Gio.Subprocess.new",
            side_effect=[workspace_process, launch_process],
        ) as launch:
            DocxWindow._open_in_writer(window)

        self.assertEqual(
            launch.call_args_list,
            [
                call(
                    ["hyprctl", "activeworkspace", "-j"],
                    Gio.SubprocessFlags.STDOUT_PIPE
                    | Gio.SubprocessFlags.STDERR_PIPE,
                ),
                call(
                    [
                        "hyprctl",
                        "dispatch",
                        "--",
                        "exec",
                        "[workspace 7 silent] libreoffice --writer /tmp/review.docx",
                    ],
                    Gio.SubprocessFlags.STDOUT_PIPE
                    | Gio.SubprocessFlags.STDERR_PIPE,
                ),
            ],
        )
        self.assertEqual(closed, [True])

    def test_keeps_preview_open_when_writer_cannot_be_started(self) -> None:
        closed: list[bool] = []
        statuses: list[str] = []
        window = _WriterWindow(statuses, closed)

        with patch(
            "see_docx.viewer.Gio.Subprocess.new",
            side_effect=[
                OSError("hyprctl not found"),
                OSError("libreoffice not found"),
            ],
        ):
            DocxWindow._open_in_writer(window)

        self.assertEqual(closed, [])
        self.assertEqual(
            statuses,
            ["Could not open in LibreOffice Writer: libreoffice not found"],
        )


if __name__ == "__main__":
    unittest.main()
