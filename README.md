# visio-mcp

An MCP server that drives **Microsoft Visio desktop** via COM automation to draw diagrams for you: flowcharts, block diagrams, and Azure/AWS architecture diagrams using the official stencils. Claude (or any MCP client) drops real Visio masters, glues dynamic connectors, runs Visio's auto-layout, and **exports a PNG of the page so the AI can see the result and iterate**.

> **Runs on Windows only** (it automates the Visio desktop app through COM). The codebase is developed and unit-tested cross-platform, but the server must run on a Windows machine with Visio installed.

## Requirements

- Windows 10/11
- Microsoft Visio desktop (any recent version; must be installed and activated)
- [Python 3.10+](https://www.python.org/downloads/) and [uv](https://docs.astral.sh/uv/getting-started/installation/)
- Claude Code (or Claude Desktop) on the same machine

## Setup (on the Windows machine)

```powershell
git clone <this repo> C:\tools\visio_mcp   # or copy the folder
cd C:\tools\visio_mcp
uv sync
```

### Azure / AWS icons

**Visio ships them built-in — no downloads needed.** Visio 16 / Microsoft 365 installs include ~130 stencils (Azure, AWS, networking, and more) under `<Office install>\root\Office16\Visio Content\<locale>\` — e.g. `AZURESTORAGE_U.VSSX`, `AWSSTORAGE_U.VSSX`. They resolve by bare filename, so `open_stencil("AZURESTORAGE_U.VSSX")` (or a fuzzy match like `open_stencil("azurestorage")`) works out of the box. `visio_status` lists them in `builtin_cloud_stencils`, and there are also dedicated **Azure Diagrams** / **AWS Diagrams** templates ([Azure](https://support.microsoft.com/en-us/visio/create-azure-diagrams-in-visio), [AWS](https://support.microsoft.com/en-us/office/create-aws-diagrams-in-visio-138206bf-d10f-4583-9f31-885ce706af49)) that `create_document` can start from.

Some masters go by different names than you might guess — the server auto-resolves common ones (`Private Endpoint` → `Private Link`, `Amazon S3` → `Bucket with Objects`, `IAM` → `Security Identity and Compliance`) and tells you when a concept is a container rather than a master (subnets, VNets, VPCs → `add_container`).

For older Visio versions or extra vendor packs, download stencils and unzip the `.vssx` files into your **My Shapes** folder (`visio_status` reports the exact path, typically `Documents\My Shapes`):

- **Azure (official icons)**: [Azure architecture icons](https://learn.microsoft.com/en-us/azure/architecture/icons/) — Microsoft now publishes SVGs; for ready-made Visio stencils use the actively maintained community packs below.
- **Azure (community `.vssx` packs)**: [Microsoft Integration and Azure Stencils Pack](https://github.com/sandroasp/Microsoft-Integration-and-Azure-Stencils-Pack-for-Visio) (Sandro Pereira, includes `MIS Azure Stencils.vssx` and many more) or [Azure-Design](https://github.com/David-Summers/Azure-Design) (David Summers).
- **AWS**: [AWS Architecture Icons](https://aws.amazon.com/architecture/icons/) — the asset package includes Visio-compatible formats; community `.vssx` conversions are also linked from that page.

`open_stencil("azure")` fuzzy-matches any `.vssx`/`.vss` file under My Shapes, and `find_masters("virtual machine")` searches inside whatever is open.

### Custom icons (no built-in master)

A few concepts have no Visio master at all — notably the newer Azure **Subnet** glyph. For these the
server keeps a **local icon folder**: drop labeled image files (`subnet.svg`, `private-endpoint.png`,
…) into it and they become first-class, reusable icons.

- `list_local_icons` reports the folder path and its contents. It defaults to `~/.visio-mcp/icons`
  and can be moved with the `VISIO_MCP_ICONS_DIR` environment variable.
- The file name without its extension is the **label**. Labels are surfaced by `find_masters`
  (marked `stencil: "(local icon)"`) and placed with `drop_shape("subnet", x, y)` — real Visio
  masters always take precedence; the local icon is the fallback before the "no master" error.
- For a one-off (or a direct image URL) you can still `import_image(source, x, y)` without saving a
  file. Icon-gallery web pages (e.g. az-icons.com) are single-page apps that return HTML rather than
  the image, so download the file (its **Download** button) or pass a direct raw image URL.

Icons are yours to supply — nothing is bundled — so respect each vendor's icon terms of use (e.g.
Microsoft's / AWS's) for anything you place there.


### Smoke test

Verifies the full pipeline end-to-end (launches Visio, builds a small styled flowchart, auto-lays it out, exports a PNG, saves/reopens the .vsdx):

```powershell
uv run python scripts/smoke_test.py
```

### Register with Claude Code

```powershell
claude mcp add visio -- uv --directory C:\tools\visio_mcp run visio-mcp
```

Then ask Claude something like: *"Draw a 5-step login flowchart in Visio and show me the result."*

## Tools

| Tool | Purpose |
|---|---|
| `visio_status` | Version, open docs/stencils, My Shapes path; launches Visio if needed |
| `create_document` | New drawing, optionally from a template (`BASFLO_U.VSTX` = basic flowchart) |
| `open_document` / `save_document` | Open/save `.vsdx` files |
| `export_page_png` | Export the page as PNG **and return the image** for visual iteration |
| `open_stencil` | Open a stencil by built-in name, path, or fuzzy My Shapes match |
| `find_masters` | Search droppable shapes across open stencils |
| `drop_shape` / `drop_shapes` | Drop masters at (x, y) inches with optional text/size |
| `update_shape` / `style_shape` / `delete_shapes` | Edit text/position/size, colors/fonts, delete |
| `connect_shapes` | Glued connector: label, arrows, right-angle/straight/curved routing, solid/dashed/dotted patterns, weight, color, optional `from_side`/`to_side` (`auto`/`top`/`bottom`/`left`/`right`) to pin endpoints for loop-backs and branches, `from_pos`/`to_pos` (0..1) for the exact point along an edge (auto-spaced when omitted), and optional `waypoints` / `label_pos` / `label_offset_in` for route control |
| `set_connector_route` | Give an existing connector explicit page-coordinate `waypoints` (native 1-D route, up to 8) and/or move its label along the route; a dynamic connector is converted while keeping endpoints, glue, arrows, style and text. `restore_auto=True` converts a fixed route back to a native dynamic connector with the same topology, static glue points, text and style |
| `route_diagnostics` | Read-only routing report: per-connector page polyline, length/detour ratio, every crossing (including at bends), every actual collinear-overlap interval, node traversal and label conflicts, with ids and exact locations; for manual routes also orthogonality and `needs_readjust`. Pass `connector_id` to focus on one connector; the summary marks `diagnostics_complete=false` / `geometry_failed_ids` when Visio geometry extraction fell back to a straight Begin-End line |
| `set_route_options` | Native page-level lane separation (`LineAdjustFrom=1`, `LineAdjustTo=2`) that keeps different connectors out of one lane without moving shapes |
| `add_container` / `container_members` | Wrap shapes in real Visio containers (VNets, VPCs, subnets, trust zones) — members move with the container |
| `drop_text` | Text-only labels (no border/fill) for titles, legends, and callouts |
| `set_page_size` | Resize the page (or fit to contents) — do this before wide architecture diagrams |
| `pages` | List/add/activate pages |
| `auto_layout` | Visio's automatic layout (flowchart top-bottom/left-right, tree, radial, circular) |
| `get_page_state` | Everything on a page: ids, masters, positions, text, connector endpoints and their glue (dynamic/static, edge, position along it, connection-point row), container membership |

**Coordinate system**: inches, origin at the page **bottom-left**, y grows upward, and drop coordinates are the shape's **center**. Rough placement is fine — finish with `auto_layout`. The default page is US Letter (8.5×11): call `set_page_size` first for wide architecture diagrams, since PNG export crops to the page bounds.

**Connector endpoints**: `connect_shapes` uses Visio's dynamic glue by default (`from_side`/`to_side` = `auto`), so Visio re-picks the best sides as shapes move. For algorithm flowcharts, pin the sides of loop-back and multi-branch edges so lines don't wrap around shapes or overlap: e.g. `connect_shapes(loop_id, cond_id, from_side="right", to_side="right")` keeps a clean loop on the right. When several connectors share one edge, each new endpoint is auto-slotted into the widest free gap with at least 0.15 in between points (the first uses the midpoint); a short, crowded edge errors rather than packing points together. Pass `from_pos`/`to_pos` to place one exactly: a fraction strictly between 0 and 1 (left/right run bottom→top, top/bottom run left→right). 0 and 1 are corners shared by two edges, so they are rejected as ambiguous — use e.g. 0.01/0.99 or the adjacent side. A pinned endpoint still follows its shape and re-routes when the page layout changes — it keeps the chosen edge and position. `get_page_state` reports each connector end's glue kind, edge, position and connection-point row.

**Route control**: when Visio's automatic router produces a shared lane or a label collision, `route_diagnostics` reports the exact connector ids, crossing points, actual overlap intervals and label rectangles. `set_route_options` turns on native lane separation for short merge trunks; `set_connector_route` (or `connect_shapes(..., waypoints=[[x, y], ...])`) pins a route to explicit page-coordinate corridors. Route-controlled connectors are native 1-D lines with static glue, normal arrowheads/text, and page-anchored geometry formulas: the waypoints stay put when nodes move or the file is reopened, while the endpoints keep following their glued connection points. Labels use Visio's `POINTALONGPATH` with `label_pos` (0..1 along the route) and `label_offset_in` (positive = left of travel). Dynamic connectors remain the default; without `waypoints` nothing changes.

**Round-trip and reliability**: `set_connector_route(connector_id, restore_auto=True)` turns a fixed route back into a native dynamic connector glued to the same connection points, preserving topology, text, arrows and line style (including the original route style when it was converted). Conversion in either direction is transactional: if a critical style cell or the text cannot be copied, the original connector is kept and the undo scope is rolled back. After a node move, a fixed route's interior vertices stay at their page positions, so the path can become non-orthogonal; `route_diagnostics` then reports `needs_readjust=true`, the offending segments and a hint to re-run `set_connector_route`. Every crossing is reported — including one that lands on a bend — and overlaps are reported as their real shared interval, not as a whole connector segment. `diagnostics_complete=false` / `geometry_failed_ids` flag connectors whose Visio geometry could not be read and were diagnosed from a straight Begin-End fallback. Use `route_diagnostics(connector_id=...)` to focus the report on one line.

## Architecture

```
Claude (stdio) → FastMCP server (asyncio; no COM imports)
                    ↓ submits callables
                 ComWorker — single STA thread owns ALL COM objects
                    ↓
                 VisioClient — the only module that touches win32com
                    ↓
                 Visio.Application
```

COM objects are apartment-threaded, so every COM call is funneled through one dedicated STA worker thread (`runtime.py`); tool bodies just `await` results. This also serializes concurrent tool calls safely.

## Development (any OS)

Unit tests run against a duck-typed fake of the Visio COM surface — no Windows needed:

```bash
uv sync
uv run pytest
```

Real-Visio connector verification (Windows + Visio desktop): `uv run python scripts/verify_connector_sides.py` — checks pinned sides, dynamic glue, node moves, auto-layout, save/reopen, and exports a PNG.

Route-control acceptance (real Visio): `uv run python scripts/verify_route_control.py` — creates a page-anchored manual route, converts a dynamic connector both ways (manual → dynamic with `restore_auto`, then back), checks the node-move `needs_readjust` diagnostic, exercises a failure rollback that must leave the original connector intact, then removes the B/C short overlaps with native lane separation and the TDMA 6.5 in overlay + label conflict with fixed corridors; writes `*-route-controlled.vsdx`/PNG and `route-control-data.json` under `D:\桌面\test\visio-mcp-regression`.

## Troubleshooting

| Symptom | Fix |
|---|---|
| "Visio is busy — close any open dialog" | A modal dialog is open in the Visio window; close it and retry |
| "Visio desktop is not installed" | Install Visio desktop; Visio for the web cannot be automated |
| "Could not open a stencil matching …" | Download the stencil pack into `Documents\My Shapes`, then retry; use `visio_status` to confirm the path |
| First run shows Visio setup/license dialogs | Start Visio manually once, dismiss the dialogs, then retry |

## Roadmap

- Agent skills encoding Azure/AWS diagram conventions (zones/containers, brand colors, icon naming)
- Optional HTTP transport to drive a Windows Visio box from another machine
- A `scripts/install_stencils.py` helper that downloads stencil packs into My Shapes
