# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Commands

```bash
# Install (editable, so local edits take effect immediately)
pip install -e .

# Run the MCP server (stdio mode — for use by Claude Desktop / Claude Code)
python -m catia_mcp

# Verify all tools load correctly (does NOT require CATIA to be running)
python test_server.py

# Lint
ruff check catia_mcp/
ruff format catia_mcp/

# Register as MCP server in Claude Code
claude mcp add catia-v5 python -- -m catia_mcp
```

## Architecture

The server bridges Claude to CATIA V5 via Windows COM automation (`win32com`).

```
Claude (stdio MCP JSON-RPC)
    └── server.py: CATIAMCPServer
            ├── Builds a tool-name → module routing table at startup
            ├── Auto-connects to CATIA on first non-connect tool call
            └── tools/
                    ├── document.py     (9 tools)   — Part/Product/document lifecycle
                    ├── sketcher.py     (11 tools)  — 2D sketch geometry & constraints
                    ├── part_design.py  (15 tools)  — 3D features (Pad, Pocket, Shaft, …)
                    ├── assembly.py     (9 tools)   — Product assembly & constraints
                    ├── measurement.py  (6 tools)   — Inertia, distance, parameters
                    ├── export.py       (4 tools)   — STEP/IGES/STL export, screenshots
                    └── scripting.py    (1 tool)    — catia_run_script bulk executor
```

All tool modules receive a shared `CATIAConnection` instance. `CATIAConnection` wraps the `win32com.client` COM object and exposes helpers (`get_active_part()`, `get_active_part_body()`, `refresh_display()`, etc.).

### Adding a new tool

1. Pick the appropriate module in `catia_mcp/tools/` (or create a new one).
2. Add an entry to `get_tool_definitions()` — every tool needs `name`, `description`, and `inputSchema`.
3. Add a `case "catia_yourname":` branch in `execute()`.
4. Implement the private method using `self.conn` for COM access.
5. If creating a new module, instantiate it in `server.py` and add it to `self._tool_modules`.
6. All tool names must start with `catia_`.

### Bulk execution — catia_run_script

For complex parts (many features, patterns, parametric geometry), use `catia_run_script` instead of individual tool calls. It executes a Python string in one MCP round-trip with `conn` (CATIAConnection) and `app` pre-injected. All tool module classes are importable inside the script. Use `print()` to return output.

```python
# Pattern for catia_run_script code argument:
from catia_mcp.tools.document import DocumentTools
from catia_mcp.tools.sketcher import SketcherTools
from catia_mcp.tools.part_design import PartDesignTools
import math

doc = DocumentTools(conn)   # conn is pre-injected
sk  = SketcherTools(conn)
pd  = PartDesignTools(conn)
# ... all operations, print() for output
```

### Sketcher session state

`SketcherTools` holds `_active_sketch` and `_active_factory` (a `Factory2D` COM object) as instance variables. These are set by `catia_create_sketch` and cleared by `catia_close_sketch`. All geometry tools (`catia_sketch_circle`, etc.) fail if called without a prior `catia_create_sketch` in the same server process session.

## COM API Gotchas

These are non-obvious issues discovered through direct CATIA V5 COM testing:

- **`part.ShapeFactory`** — the correct accessor for Part Design features. `body.ShapeFactory` does not exist.
- **`Part.Name` is read-only** — setting `doc.Part.Name = x` raises a COM error. The part name is set by CATIA automatically.
- **Sketch geometry cannot be deleted via COM** — `sketch.Remove(element)` silently does nothing. To "move" Part Design geometry, close the broken document without saving and rebuild the sketch at the new position.
- **`part.Update()`** fails when a sketch has overlapping/invalid geometry. Close the document and start fresh rather than trying to recover.
- **Pocket direction** — when a Pad extrudes in +Z from the XY plane, the Pocket from the same XY plane must use `"direction": "reverse"` to cut into the solid; the default cuts in -Z (into empty space). When the pad orientation is uncertain, make the cut **direction-independent** by also extending the second limit: `pocket.SecondLimit.Dimension.Value = depth` (cuts both ways through the stock regardless of which side the solid is on). The same `FirstLimit`/`SecondLimit` `.Dimension.Value` properties work on Pads to grow a feature past a face on each side (e.g. liner overhang).
- **All coordinates and distances are in millimeters** — CATIA V5 COM uses mm natively.
- **`Factory2D.CreateClosedCircle(cx, cy, radius)`** — use this for full circles, not `CreateCircle`.
- **Two concentric circles in one sketch** — CATIA interprets the inner circle as a hole when padded, producing an annular ring automatically.
- **Multiple closed profiles in one Pocket sketch** — all profiles are cut simultaneously in one Pocket feature. Use this to cut many slots in a single operation instead of a circular pattern.
- **COM thread initialisation** — `pythoncom.CoInitialize()` is called once in `CATIAConnection.connect()`. Do not call it again in tool methods.
- **`pyproject.toml` build backend** — must be `"setuptools.build_meta"`, not `"setuptools.backends._legacy:_Backend"` (the latter causes `BackendUnavailable` with pip's bundled pyproject-hooks).
- **Multiple bodies in one Part** — `part.Bodies.Add()` creates a new body; set `part.InWorkObject = body` before adding features so they land in it. Bodies stay as separate solids (not unioned), which is how to represent distinct materials in one Part (e.g. steel core / insulation / wedges). Delete a body or feature with `Selection.Add(obj)` then `Selection.Delete()`.
- **Colouring a body/feature** — `sel = doc.Selection; sel.Clear(); sel.Add(body); sel.VisProperties.SetRealColor(r, g, b, 0); sel.Clear()` (RGB 0–255).
- **Offset sketch plane** — `plane = part.HybridShapeFactory.AddNewPlaneOffset(ref, dist_mm, False)`, then `body.InsertHybridShape(plane)` and `part.UpdateObject(plane)` before `body.Sketches.Add(part.CreateReferenceFromObject(plane))`.
- **Screenshots are TIFF, not PNG** — `app.ActiveWindow.ActiveViewer.CaptureToFile(3, path)` writes a **TIFF** byte stream (`II*` header) regardless of the file extension. Save as `.tif` and convert with Pillow (`Image.open(p).save(p2)`) if a PNG is needed.
- **`GetMeasurable(...).Volume` returns 0** — the SPAWorkbench measurable reports `0` for solid volume/area on this build; do not rely on it for verification. Use a screenshot or a feature/parameter check instead.

## Known COM limitations (do NOT attempt via this bridge)

- **Multi-Sections Solid / Loft / RemovedLoft are not drivable.** `ShapeFactory.AddNewLoft()` returns an object whose `AddSectionToLoft` method is **unreachable through late binding** (the object exposes no usable type info, so `GetIDsOfNames` fails). The only workaround — generating a `makepy`/gen_py early-binding module — produces an **incomplete** binding for CATIA's typelibs (`AddSectionToLoft` is missing from the generated `Loft` class) **and**, while the gen_py cache exists on disk, it makes win32com wrap *all* returned objects as base interfaces, breaking late-bound access to every other factory (`ShapeFactory`, `HybridShapeFactory`, etc.). If you ever generate it, delete it afterwards: `shutil.rmtree(win32com.client.gencache.GetGeneratePath())`. **Practical consequence:** lofted operations such as **slot skew** cannot be scripted here — do them in the CATIA UI, or approximate with stacked rotated slices (pads/pockets), which works fine.

## Connection Lifecycle

`CATIAConnection.connect()` tries `win32com.client.GetActiveObject("CATIA.Application")` first (attach to running instance), then falls back to `win32com.client.Dispatch("CATIA.Application")` (launch new instance). The server auto-connects on the first tool call that is not `catia_connect` or `catia_disconnect`.

## Logging

The server logs to both `catia_mcp.log` (in the working directory) and stderr. Tool call names and truncated results are logged at INFO level; exceptions at ERROR with full tracebacks.
