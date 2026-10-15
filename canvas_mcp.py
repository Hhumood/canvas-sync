"""
canvas_mcp.py - a tiny MCP server for the Work Week dashboard.

It has one tool, `sync`. When the dashboard calls it, this server runs the Canvas
code in canvas_fetch.py and hands the result straight back to the dashboard.

How it fits together:
  Dashboard (in the Claude app)  --calls tool "sync"-->  this server (on your Mac)
  this server  --runs-->  canvas_fetch.build_export()  --asks-->  Canvas API
  this server  --returns the export (JSON)-->  dashboard merges it into your planner

The Canvas token never leaves your Mac: canvas_fetch.py reads it from .env here,
and only the assignments and announcements go back to the dashboard.

The Claude app starts this file by itself (see claude_desktop_config.json).
Nothing to run by hand.
"""

import json
import subprocess
import sys
from pathlib import Path

from mcp.server.fastmcp import FastMCP
from mcp.types import ToolAnnotations

HERE = Path(__file__).resolve().parent
MARK = "@@CANVAS_EXPORT@@"

# The Canvas code runs in a separate Python process. An MCP server talks to the app
# through its standard output, so anything canvas_fetch prints must not land there.
# The child prints the export after a marker, and we read only what follows it.
CHILD = (
    "import json, canvas_fetch as f\n"
    "if not f.cs.TOKEN: raise SystemExit('No Canvas token found. Save it in the .env file in canvas-sync.')\n"
    f"print({MARK!r} + json.dumps(f.build_export(), ensure_ascii=False))\n"
)

server = FastMCP("canvas-sync")


@server.tool(
    annotations=ToolAnnotations(title="Sync Canvas", readOnlyHint=True, openWorldHint=True),
    structured_output=False,
)
def sync() -> str:
    """Pull the latest assignments and announcements from Canvas and return them as JSON.
    Reads Canvas only. Nothing on Canvas is changed and no file is written."""
    try:
        run = subprocess.run(
            [sys.executable, "-B", "-c", CHILD],
            cwd=HERE, capture_output=True, text=True, timeout=110,
        )
    except subprocess.TimeoutExpired:
        raise RuntimeError("Canvas took longer than 110 seconds to answer. Try again in a minute.")
    out = run.stdout or ""
    if run.returncode != 0 or MARK not in out:
        lines = (run.stderr or out).strip().splitlines()
        raise RuntimeError(lines[-1][:300] if lines else "The Canvas script stopped without saying why.")
    return out.split(MARK, 1)[1].strip()


if __name__ == "__main__":
    server.run()
