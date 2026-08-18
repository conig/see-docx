"""Small Sway IPC helpers shared by the graphical smoke tests."""

from __future__ import annotations

import json
import os
from pathlib import Path
import select
import subprocess
from collections.abc import Iterator


WORKSPACE = "15"


def capture_failure_screenshot() -> None:
    """Capture the still-mapped test surface before a smoke test closes it."""

    destination = os.environ.get("SEE_DOCX_GUI_FAILURE_SCREENSHOT")
    if destination:
        subprocess.run(
            ["grim", "-c", destination],
            check=False,
            stdout=subprocess.DEVNULL,
            stderr=subprocess.DEVNULL,
        )


class DesktopInput:
    """Drive compositor-native input in the isolated headless Sway session."""

    def __init__(self, directory: str) -> None:
        output = next(
            (
                candidate
                for candidate in sway_json("get_outputs")
                if candidate.get("active")
            ),
            None,
        )
        if output is None:
            raise RuntimeError("headless Sway has no active output")
        self._extent = (output["rect"]["width"], output["rect"]["height"])
        self._cursor = (0, 0)
        self._keyboard_processes: list[subprocess.Popen[str]] = []
        build = Path(directory) / "virtual-pointer"
        header = Path(directory) / (
            "wlr-virtual-pointer-unstable-v1-client-protocol.h"
        )
        protocol_source = Path(directory) / (
            "wlr-virtual-pointer-unstable-v1-protocol.c"
        )
        tests = Path(__file__).resolve().parent
        protocol = tests / "protocols" / "wlr-virtual-pointer-unstable-v1.xml"
        subprocess.run(
            ["wayland-scanner", "client-header", protocol, header],
            check=True,
            capture_output=True,
            text=True,
        )
        subprocess.run(
            ["wayland-scanner", "private-code", protocol, protocol_source],
            check=True,
            capture_output=True,
            text=True,
        )
        subprocess.run(
            [
                "gcc",
                "-std=c11",
                "-Wall",
                "-Wextra",
                "-Werror",
                f"-I{directory}",
                tests / "wayland_virtual_pointer.c",
                protocol_source,
                "-o",
                build,
                "-lwayland-client",
                "-lm",
            ],
            check=True,
            capture_output=True,
            text=True,
        )
        self._pointer = subprocess.Popen(
            [build],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            bufsize=1,
        )
        if self._pointer.stdout is None:
            raise RuntimeError("virtual pointer has no readiness stream")
        ready, _writeable, _exceptional = select.select(
            [self._pointer.stdout], [], [], 3.0
        )
        if not ready or self._pointer.stdout.readline().strip() != "READY":
            error = (
                self._pointer.stderr.read().strip()
                if self._pointer.stderr is not None
                else ""
            )
            raise RuntimeError(f"Wayland virtual pointer did not start: {error}")

    def move_cursor(self, x: int, y: int) -> None:
        self._cursor = (x, y)
        self._pointer_command(
            f"motion_absolute {x} {y} {self._extent[0]} {self._extent[1]}"
        )

    def cursor_position(self) -> tuple[int, int]:
        return self._cursor

    def _pointer_command(self, command: str) -> None:
        if self._pointer.poll() is not None or self._pointer.stdin is None:
            error = (
                self._pointer.stderr.read().strip()
                if self._pointer.stderr is not None
                else ""
            )
            raise RuntimeError(f"Wayland virtual pointer stopped: {error}")
        self._pointer.stdin.write(f"{command}\n")
        self._pointer.stdin.flush()

    def left_button(self, *, pressed: bool) -> None:
        self._pointer_command(f"button 272 {1 if pressed else 0}")

    def scroll_down(self) -> None:
        self._pointer_command("scroll 15 1")

    def _keyboard(self, *arguments: str) -> None:
        # Keep the virtual keyboard alive while GTK's main loop processes the
        # new wl_seat capability and binds wl_keyboard. Running wtype to
        # completion inside a GTK callback creates and removes the device
        # before the client can bind it, so no key event can be delivered.
        self._keyboard_processes.append(
            subprocess.Popen(
                ["wtype", "-s", "100", *arguments, "-s", "100"],
                stdout=subprocess.PIPE,
                stderr=subprocess.PIPE,
                text=True,
            )
        )

    def type_text(self, value: str) -> None:
        self._keyboard(value)

    def key(self, value: str) -> None:
        self._keyboard("-k", value)

    def close(self) -> None:
        for process in self._keyboard_processes:
            try:
                process.wait(timeout=2)
            except subprocess.TimeoutExpired:
                process.terminate()
                process.wait(timeout=2)
        if self._pointer.stdin is not None:
            try:
                self._pointer.stdin.write("quit\n")
                self._pointer.stdin.flush()
            except BrokenPipeError:
                pass
        try:
            self._pointer.wait(timeout=2)
        except subprocess.TimeoutExpired:
            self._pointer.terminate()
            try:
                self._pointer.wait(timeout=2)
            except subprocess.TimeoutExpired:
                self._pointer.kill()
                self._pointer.wait(timeout=2)


def sway_json(message_type: str) -> object:
    """Return one JSON IPC response from the isolated test compositor."""

    return json.loads(
        subprocess.run(
            ["swaymsg", "-t", message_type, "-r"],
            check=True,
            capture_output=True,
            text=True,
        ).stdout
    )


def sway_command(command: str) -> None:
    """Run one Sway command and reject unsuccessful IPC results."""

    response = json.loads(
        subprocess.run(
            ["swaymsg", "-r", command],
            check=True,
            capture_output=True,
            text=True,
        ).stdout
    )
    if not response or not all(result.get("success") for result in response):
        raise RuntimeError(f"Sway rejected {command!r}: {response!r}")


def _nodes(
    node: dict[str, object], workspace: str | None = None
) -> Iterator[tuple[dict[str, object], str | None]]:
    if node.get("type") == "workspace":
        workspace = str(node["name"])
    yield node, workspace
    for child in [*node.get("nodes", []), *node.get("floating_nodes", [])]:
        yield from _nodes(child, workspace)


def clients(
    *, app_id: str | None = None, pid: int | None = None
) -> list[dict[str, object]]:
    """Return normalized client geometry plus the containing workspace."""

    root = sway_json("get_tree")
    matches: list[dict[str, object]] = []
    for node, workspace in _nodes(root):
        if node.get("type") not in {"con", "floating_con"}:
            continue
        if app_id is not None and node.get("app_id") != app_id:
            continue
        if pid is not None and node.get("pid") != pid:
            continue
        if node.get("app_id") is None and node.get("window") is None:
            continue
        rect = node["rect"]
        matches.append(
            {
                "id": node["id"],
                "app_id": node.get("app_id"),
                "pid": node.get("pid"),
                "at": [rect["x"], rect["y"]],
                "size": [rect["width"], rect["height"]],
                "workspace": {"id": int(workspace), "name": workspace},
            }
        )
    return matches


def smoke_client(app_id: str) -> dict[str, object] | None:
    """Find the current process's one smoke-test surface."""

    return next(iter(clients(app_id=app_id, pid=os.getpid())), None)


def focus_client(client: dict[str, object]) -> None:
    sway_command(f"[con_id={client['id']}] focus")


def focus_workspace(workspace: str = WORKSPACE) -> None:
    sway_command(f"workspace number {workspace}")
