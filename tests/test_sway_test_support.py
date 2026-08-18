from __future__ import annotations

from types import SimpleNamespace
import unittest
from unittest.mock import patch

import sway_test_support


class SwayTestSupportTests(unittest.TestCase):
    def test_clients_normalize_geometry_and_containing_workspace(self) -> None:
        tree = {
            "type": "root",
            "nodes": [
                {
                    "type": "output",
                    "nodes": [
                        {
                            "type": "workspace",
                            "name": "15",
                            "nodes": [
                                {
                                    "type": "con",
                                    "id": 42,
                                    "app_id": "codex-smoke-test",
                                    "pid": 123,
                                    "rect": {
                                        "x": 11,
                                        "y": 12,
                                        "width": 1440,
                                        "height": 900,
                                    },
                                    "nodes": [],
                                    "floating_nodes": [],
                                }
                            ],
                            "floating_nodes": [],
                        }
                    ],
                    "floating_nodes": [],
                }
            ],
            "floating_nodes": [],
        }
        with patch.object(sway_test_support, "sway_json", return_value=tree):
            self.assertEqual(
                sway_test_support.clients(
                    app_id="codex-smoke-test", pid=123
                ),
                [
                    {
                        "id": 42,
                        "app_id": "codex-smoke-test",
                        "pid": 123,
                        "at": [11, 12],
                        "size": [1440, 900],
                        "workspace": {"id": 15, "name": "15"},
                    }
                ],
            )

    def test_sway_command_rejects_an_ipc_level_failure(self) -> None:
        completed = SimpleNamespace(
            stdout='[{"success":false,"error":"no matching node"}]'
        )
        with patch.object(
            sway_test_support.subprocess, "run", return_value=completed
        ):
            with self.assertRaisesRegex(RuntimeError, "no matching node"):
                sway_test_support.sway_command("[con_id=9] focus")


if __name__ == "__main__":
    unittest.main()
