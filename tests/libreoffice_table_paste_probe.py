#!/usr/bin/env python3
"""Paste See DOCX's live clipboard into a blank disposable Writer.

The probe intentionally does not open the embedded ODF clipboard bytes and
copy them again with Writer.  Doing that replaces the clipboard owner and only
tests Writer-to-Writer copy/paste, which cannot detect an invalid See DOCX
clipboard representation.
"""

from __future__ import annotations

import argparse
import json
import os
from pathlib import Path
import subprocess
import tempfile
import time

import uno

from sway_test_support import WORKSPACE, clients, focus_client, focus_workspace

WRITER_APP_ID = "libreoffice-writer"


def main() -> int:
    parser = argparse.ArgumentParser()
    parser.add_argument("mode", choices=("row", "column"))
    arguments = parser.parse_args()

    with tempfile.TemporaryDirectory(
        prefix="see-docx-libreoffice-table-paste-"
    ) as directory:
        pipe_name = f"see_docx_table_paste_{os.getpid()}"
        profile_uri = (Path(directory) / "profile").as_uri()
        existing_ids = {client["id"] for client in clients(app_id=WRITER_APP_ID)}
        process: subprocess.Popen[bytes] | None = None
        desktop = None
        blank_document = None
        try:
            focus_workspace(WORKSPACE)
            process = subprocess.Popen(
                [
                    "env",
                    "SAL_USE_VCLPLUGIN=gtk3",
                    "libreoffice",
                    f"-env:UserInstallation={profile_uri}",
                    "--nodefault",
                    "--nologo",
                    "--norestore",
                    "--nolockcheck",
                    f"--accept=pipe,name={pipe_name};urp;StarOffice.ComponentContext",
                ],
                stdout=subprocess.DEVNULL,
                stderr=subprocess.DEVNULL,
            )
            local_context = uno.getComponentContext()
            resolver = local_context.ServiceManager.createInstanceWithContext(
                "com.sun.star.bridge.UnoUrlResolver", local_context
            )
            context = None
            for _attempt in range(200):
                try:
                    context = resolver.resolve(
                        f"uno:pipe,name={pipe_name};urp;StarOffice.ComponentContext"
                    )
                    break
                except Exception:
                    time.sleep(0.05)
            if context is None:
                raise RuntimeError("could not connect to the disposable Writer")

            services = context.ServiceManager
            desktop = services.createInstanceWithContext(
                "com.sun.star.frame.Desktop", context
            )
            blank_document = desktop.loadComponentFromURL(
                "private:factory/swriter", "_blank", 0, ()
            )
            if blank_document is None:
                raise RuntimeError("could not create the disposable Writer document")
            writer_clients: list[dict[str, object]] = []
            for _attempt in range(200):
                writer_clients = [
                    client
                    for client in clients(app_id=WRITER_APP_ID)
                    if client["id"] not in existing_ids
                ]
                if writer_clients:
                    break
                time.sleep(0.05)
            if not writer_clients:
                raise RuntimeError("the disposable Writer window did not map")
            client = writer_clients[0]
            if client["workspace"]["id"] != int(WORKSPACE):
                raise RuntimeError(
                    f"Writer mapped outside workspace 15: {client['workspace']}"
                )

            dispatch = services.createInstanceWithContext(
                "com.sun.star.frame.DispatchHelper", context
            )
            focus_workspace(WORKSPACE)
            focus_client(client)
            dispatch.executeDispatch(
                blank_document.CurrentController.Frame, ".uno:Paste", "", 0, ()
            )
            time.sleep(0.5)
            pasted_tables = blank_document.TextTables
            blank_tables = []
            for table_index in range(pasted_tables.getCount()):
                pasted_table = pasted_tables.getByIndex(table_index)
                blank_tables.append(
                    [
                        [
                            pasted_table.getCellByName(f"{column}{row}").String
                            for column in "ABC"[: pasted_table.Columns.getCount()]
                        ]
                        for row in range(1, pasted_table.Rows.getCount() + 1)
                    ]
                )
            blank_text = blank_document.Text.String
            print(
                json.dumps(
                    {
                        "blank_tables": blank_tables,
                        "blank_text": blank_text,
                    }
                )
            )
        finally:
            if blank_document is not None:
                try:
                    blank_document.close(True)
                except Exception:
                    pass
            if desktop is not None:
                try:
                    desktop.terminate()
                except Exception:
                    pass
            if process is not None and process.poll() is None:
                process.terminate()
                try:
                    process.wait(timeout=3)
                except subprocess.TimeoutExpired:
                    process.kill()
                    process.wait(timeout=3)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
