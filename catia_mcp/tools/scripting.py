"""Scripting tool for CATIA V5.

Executes a Python script in one MCP call with conn (CATIAConnection) pre-injected.
Use for bulk geometry operations to avoid per-call MCP JSON-RPC overhead.
"""

from __future__ import annotations

import io
import sys
import traceback
from typing import Any

from catia_mcp.connection import CATIAConnection


class ScriptingTools:
    """Bulk-execution tool: run a full Python script inside the server process."""

    def __init__(self, connection: CATIAConnection) -> None:
        self.conn = connection

    def get_tool_definitions(self) -> list[dict[str, Any]]:
        return [
            {
                "name": "catia_run_script",
                "description": (
                    "Execute a Python script inside the CATIA server process in a single MCP call. "
                    "Use this for bulk geometry operations (complex parts, patterns, assemblies) "
                    "to avoid per-operation round-trip overhead. "
                    "Pre-injected variables: conn (CATIAConnection), app (CATIA Application COM object). "
                    "All catia_mcp tool classes are importable. Use print() to return output."
                ),
                "inputSchema": {
                    "type": "object",
                    "properties": {
                        "code": {
                            "type": "string",
                            "description": (
                                "Python code to execute. "
                                "conn and app are available as globals. "
                                "Use print() for results — all stdout is returned."
                            ),
                        },
                    },
                    "required": ["code"],
                },
            }
        ]

    def execute(self, tool_name: str, arguments: dict[str, Any]) -> str:
        match tool_name:
            case "catia_run_script":
                return self._run_script(arguments["code"])
            case _:
                raise ValueError(f"Unknown scripting tool: {tool_name}")

    def _run_script(self, code: str) -> str:
        self.conn.ensure_connected()

        buf = io.StringIO()
        old_stdout = sys.stdout
        sys.stdout = buf
        try:
            namespace: dict[str, Any] = {
                "conn": self.conn,
                "app": self.conn.app,
            }
            exec(code, namespace)  # noqa: S102
            output = buf.getvalue()
            return output.strip() if output.strip() else "Script executed successfully (no output)"
        except Exception:
            return f"Script error:\n{traceback.format_exc()}"
        finally:
            sys.stdout = old_stdout
