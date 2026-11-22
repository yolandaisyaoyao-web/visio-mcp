"""visio-mcp: MCP server that drives Microsoft Visio desktop via COM.

Tool bodies never touch COM directly: every call is submitted to the
ComWorker's single STA thread, which owns the VisioClient and all COM
objects (see runtime.py for why).
"""

from __future__ import annotations

import asyncio
import functools
import os
import tempfile
import time
import urllib.parse
import urllib.request
from typing import Literal, Optional

from mcp.server.fastmcp import FastMCP, Image
from mcp.server.fastmcp.exceptions import ToolError

from .errors import VisioMcpError
from .models import DropSpec
from .runtime import ComWorker
from .visio_client import VisioClient

mcp = FastMCP(
    "visio",
    instructions=(
        "Draw diagrams in a live Microsoft Visio desktop instance (Windows). "
        "Typical workflow: visio_status -> create_document (e.g. template "
        "'BASFLO_U.VSTX' for flowcharts) -> open_stencil / find_masters to "
        "discover shapes -> drop_shapes -> connect_shapes -> auto_layout -> "
        "export_page_png to SEE the result and iterate -> save_document. "
        "Coordinate system: units are INCHES, origin at the page BOTTOM-LEFT, "
        "y grows UPWARD, and the drop point (x, y) is the shape's CENTER. "
        "Rough placement is fine when you finish with auto_layout. For wide "
        "architecture diagrams, call set_page_size FIRST (default page is "
        "8.5x11 and the PNG export crops to it). Group shapes into zones "
        "(VNet/VPC/subnet/subscription/trust boundary) with add_container — "
        "innermost containers first — and badge a zone with its header icon "
        "via badge_container (deterministic corner placement; the badge moves "
        "with the zone). Use drop_text for titles and legends, and dashed "
        "connectors (line_pattern) for control-plane/auth flows. Pin "
        "connector endpoints with connect_shapes from_side/to_side "
        "(auto/top/bottom/left/right) for loop-backs and branches — e.g. "
        "loop back with right->right so the line doesn't wrap around shapes. "
        "Several connectors on one edge are auto-spaced into free slots; use "
        "from_pos/to_pos (0..1) to place one exactly. get_page_state reports "
        "each connector end's glue kind, edge, position and point row. "
        "When Visio's router shares a lane or a label collides, call "
        "route_diagnostics (read-only): it reports the polyline, crossings, "
        "overlaps, node traversal and label conflicts with ids and points. "
        "Fix short merge trunks with set_route_options(separate_lanes=True); "
        "pin exact corridors with connect_shapes(waypoints=...) or "
        "set_connector_route on an existing connector (native 1-D line, "
        "endpoints stay glued). Place labels with POINTALONGPATH via "
        "label_pos/label_offset_in. Dynamic routing remains the default. "
        "For Azure/"
        "AWS icons, Visio ships built-in cloud stencils — visio_status lists "
        "them in builtin_cloud_stencils and they open by bare filename (e.g. "
        "open_stencil('AZURESTORAGE_U.VSSX')); extra vendor packs can go in "
        "the My Shapes folder. Search open stencils with find_masters. For an "
        "icon Visio has no master for (e.g. the Azure 'Subnet' glyph), save a "
        "labeled image (subnet.svg) in the local icon folder (list_local_icons "
        "shows the path) so drop_shape('subnet', ...) places it, or insert one "
        "on the fly with import_image (a local file path or a direct image "
        "URL)."
    ),
)

_worker = ComWorker()
_client = VisioClient()
_export_counter = 0


def _tool_errors(fn):
    """Convert VisioMcpError (and unexpected errors) into clean ToolErrors."""

    @functools.wraps(fn)
    async def wrapper(*args, **kwargs):
        try:
            return await fn(*args, **kwargs)
        except (VisioMcpError, ValueError) as exc:
            raise ToolError(str(exc)) from exc

    return wrapper


_IMAGE_EXTS = {".svg", ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".emf", ".wmf"}
_CTYPE_EXT = {
    "image/svg+xml": ".svg",
    "image/png": ".png",
    "image/jpeg": ".jpg",
    "image/gif": ".gif",
    "image/bmp": ".bmp",
}


def _download_image_to_temp(url: str) -> str:
    """Download a direct image URL to a temp file for Visio to Import.

    Guards against the common failure where an icon *gallery* URL returns an
    HTML single-page app instead of the raw image bytes.
    """
    req = urllib.request.Request(url, headers={"User-Agent": "visio-mcp"})
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            ctype = (resp.headers.get("Content-Type") or "").split(";")[0].strip().lower()
            data = resp.read()
    except Exception as exc:
        raise VisioMcpError(f"Failed to download image from {url!r}: {exc}") from exc
    head = data[:64].lstrip().lower()
    if ctype == "text/html" or head.startswith(b"<!doctype html") or head.startswith(b"<html"):
        raise VisioMcpError(
            f"{url!r} returned an HTML page, not an image. Icon galleries like "
            "az-icons.com are single-page apps whose /icon URLs serve HTML — "
            "use the site's Download button and pass the saved file path, or a "
            "direct raw image URL (one that ends in .svg/.png)."
        )
    ext = os.path.splitext(urllib.parse.urlparse(url).path)[1].lower()
    if ext not in _IMAGE_EXTS:
        ext = _CTYPE_EXT.get(ctype, ".img")
    fd, tmp = tempfile.mkstemp(prefix="visio_mcp_icon_", suffix=ext)
    os.close(fd)
    with open(tmp, "wb") as fh:
        fh.write(data)
    return tmp


@mcp.tool()
@_tool_errors
async def visio_status() -> dict:
    """Report Visio state: version, open documents and stencils, active page,
    the My Shapes folder path, and the built-in stencil library — including
    builtin_cloud_stencils, the Azure/AWS stencil files Visio ships with
    (openable by bare filename, no download needed). Launches Visio if it is
    not already running. Call this first."""
    return await _worker.run(_client.status)


@mcp.tool()
@_tool_errors
async def create_document(
    template: Optional[str] = None,
    measurement: Literal["us", "metric"] = "us",
) -> dict:
    """Create a new drawing document in Visio.

    Args:
        template: Visio template name or absolute path. Examples:
            'BASFLO_U.VSTX' (Basic Flowchart, US units) or bare 'BASFLO'
            (units suffix added from `measurement`). Omit for a blank drawing.
        measurement: 'us' (inches) or 'metric' — used when template has no
            _U/_M suffix.
    """
    return await _worker.run(_client.create_document, template, measurement)


@mcp.tool()
@_tool_errors
async def open_document(path: str) -> dict:
    """Open an existing Visio file (.vsdx/.vsd) for viewing or editing."""
    return await _worker.run(_client.open_document, path)


@mcp.tool()
@_tool_errors
async def save_document(path: Optional[str] = None) -> dict:
    """Save the active drawing. Pass an absolute path ending in .vsdx the
    first time; afterwards you can omit it to save in place."""
    return await _worker.run(_client.save_document, path)


@mcp.tool()
@_tool_errors
async def export_page_png(
    path: Optional[str] = None,
    page: Optional[str] = None,
) -> list:
    """Export a page as PNG and RETURN THE IMAGE so you can visually inspect
    the diagram. Use this after dropping/connecting/layout to verify the
    result and iterate.

    Args:
        path: Output .png path; omit for an auto-named temp file.
        page: Page name; defaults to the active page.
    """
    global _export_counter
    if path is None:
        _export_counter += 1
        path = os.path.join(
            tempfile.gettempdir(),
            f"visio_mcp_{os.getpid()}_{_export_counter}_{int(time.time())}.png",
        )
    result = await _worker.run(_client.export_page_png, path, page)
    return [f"Exported page {result['page']!r} to {result['path']}", Image(path=result["path"])]


@mcp.tool()
@_tool_errors
async def open_stencil(name_or_path: str) -> dict:
    """Open a stencil (shape library) so its masters can be dropped.

    Accepts a built-in name ('BASFLO_U.VSSX' basic flowchart, 'PERIPH_U.VSSX'
    peripherals...), an absolute path, or a fuzzy name matched against files
    in the My Shapes folder (where downloaded Azure/AWS stencil packs live).
    """
    return await _worker.run(_client.open_stencil, name_or_path)


@mcp.tool()
@_tool_errors
async def find_masters(query: Optional[str] = None, stencil: Optional[str] = None) -> dict:
    """Search masters (droppable shapes) in the open stencils by name
    substring. Omit `query` to list everything (capped at 100). Matching
    custom icons from the local icon folder are included too, marked with
    stencil '(local icon)'."""
    return await _worker.run(_client.find_masters, query, stencil)


@mcp.tool()
@_tool_errors
async def list_local_icons() -> dict:
    """List the local custom-icon folder and the labeled image files in it.

    Users manually download icons (e.g. from az-icons.com) and save them into
    this folder; each file's name without its extension is its label. Labels
    are discoverable via find_masters and can be placed with
    drop_shape('<label>', x, y) when Visio has no built-in master. The folder
    defaults to ~/.visio-mcp/icons and moves with the VISIO_MCP_ICONS_DIR
    environment variable."""
    return await _worker.run(_client.list_local_icons)


@mcp.tool()
@_tool_errors
async def drop_shape(
    master: str,
    x: float,
    y: float,
    stencil: Optional[str] = None,
    text: Optional[str] = None,
    width_in: Optional[float] = None,
    height_in: Optional[float] = None,
    page: Optional[str] = None,
) -> dict:
    """Drop one master onto the page. (x, y) is the shape CENTER in inches
    from the bottom-left. Returns the shape_id used by other tools."""
    spec = DropSpec(
        master=master, x=x, y=y, stencil=stencil, text=text,
        width_in=width_in, height_in=height_in,
    ).model_dump()
    result = await _worker.run(_client.drop_shapes, [spec], page)
    return result["shapes"][0]


@mcp.tool()
@_tool_errors
async def drop_shapes(shapes: list[DropSpec], page: Optional[str] = None) -> dict:
    """Drop many masters in one call (preferred when building a diagram).
    Executed in a single undo scope. Returns a shape_id per item, in order."""
    specs = [s.model_dump() for s in shapes]
    return await _worker.run(_client.drop_shapes, specs, page)


@mcp.tool()
@_tool_errors
async def import_image(
    source: str,
    x: float,
    y: float,
    width_in: Optional[float] = None,
    height_in: Optional[float] = None,
    page: Optional[str] = None,
) -> dict:
    """Insert an external image (SVG, PNG, JPG, EMF, ...) onto the page as a
    shape — for custom icons that have NO Visio master (e.g. the newer Azure
    'Subnet' glyph) or vendor logos.

    Args:
        source: a LOCAL file path OR a direct http(s) IMAGE URL. A URL must
            point at the raw image bytes; icon *web pages* (single-page-app
            galleries such as az-icons.com) return HTML and are rejected with
            guidance — download the file in the browser and pass its path.
        x, y: shape CENTER in inches from the bottom-left.
        width_in / height_in: size in inches. Give just one to scale the other
            automatically and preserve the image's aspect ratio; omit both to
            keep the image's native size.

    Prefer built-in masters (drop_shape) when Visio has the icon; use this for
    the gaps. Returns the shape_id.
    """
    path = source
    if source.strip().lower().startswith(("http://", "https://")):
        loop = asyncio.get_running_loop()
        path = await loop.run_in_executor(None, _download_image_to_temp, source.strip())
    return await _worker.run(
        _client.import_image, path, x, y, width_in, height_in, page
    )


@mcp.tool()
@_tool_errors
async def update_shape(
    shape_id: int,
    text: Optional[str] = None,
    x: Optional[float] = None,
    y: Optional[float] = None,
    width_in: Optional[float] = None,
    height_in: Optional[float] = None,
    page: Optional[str] = None,
) -> dict:
    """Change a shape's text, position (center, inches), or size. Only the
    provided fields are changed."""
    return await _worker.run(
        _client.update_shape, shape_id, page, text, x, y, width_in, height_in
    )


@mcp.tool()
@_tool_errors
async def style_shape(
    shape_id: int,
    fill_color: Optional[str] = None,
    line_color: Optional[str] = None,
    text_color: Optional[str] = None,
    line_weight_pt: Optional[float] = None,
    font_size_pt: Optional[float] = None,
    bold: Optional[bool] = None,
    line_pattern: Optional[Literal["solid", "dashed", "dotted", "dash_dot"]] = None,
    page: Optional[str] = None,
) -> dict:
    """Style a shape, connector, or container. Colors are hex like '#0078D4'
    (Azure blue) or '#FF9900' (AWS orange). Only the provided fields are
    changed. Works on containers too: theme-guarded cells are force-
    overridden, so recoloring a VNet/subscription boundary is supported."""
    return await _worker.run(
        _client.style_shape, shape_id, page, fill_color, line_color,
        text_color, line_weight_pt, font_size_pt, bold, line_pattern,
    )


@mcp.tool()
@_tool_errors
async def delete_shapes(shape_ids: list[int], page: Optional[str] = None) -> dict:
    """Delete shapes (or connectors) by id."""
    return await _worker.run(_client.delete_shapes, shape_ids, page)


@mcp.tool()
@_tool_errors
async def connect_shapes(
    from_id: int,
    to_id: int,
    label: Optional[str] = None,
    route: Literal["right_angle", "straight", "curved"] = "right_angle",
    end_arrow: bool = True,
    begin_arrow: bool = False,
    line_pattern: Literal["solid", "dashed", "dotted", "dash_dot"] = "solid",
    line_weight_pt: Optional[float] = None,
    line_color: Optional[str] = None,
    page: Optional[str] = None,
    from_side: Literal["auto", "top", "bottom", "left", "right"] = "auto",
    to_side: Literal["auto", "top", "bottom", "left", "right"] = "auto",
    from_pos: Optional[float] = None,
    to_pos: Optional[float] = None,
    waypoints: Optional[list[list[float]]] = None,
    label_pos: Optional[float] = None,
    label_offset_in: Optional[float] = None,
) -> dict:
    """Connect two shapes with a connector glued to both. With the default
    from_side='auto' / to_side='auto' the connector uses Visio's dynamic glue:
    Visio picks the best sides and re-routes when shapes move. Pass a side to
    pin an endpoint deterministically — use it for loop-backs, decision
    branches and parallel edges so lines don't wrap around shapes or overlap.

    Args:
        label: Optional text on the connector (e.g. 'Yes' / 'No').
        route: 'right_angle' (default), 'straight', or 'curved'.
        line_pattern: 'dashed'/'dotted' are the convention for control-plane,
            auth, or reference flows in architecture diagrams.
        line_weight_pt: Thicker lines (e.g. 2) suit peering/trust boundaries.
        line_color: Hex like '#808080'.
        from_side: 'auto' (dynamic glue), or 'top'/'bottom'/'left'/'right' to
            attach the connector's begin to an edge of that shape. A pinned
            endpoint follows its shape when it moves but keeps the chosen
            edge and position — typical for a loop-back (right -> right) or
            a decision's Yes/No branches.
        to_side: Same for the arrowhead end at to_id.
        from_pos: Optional position along from_side, strictly between 0 and 1
            (left/right edges run bottom -> top, top/bottom edges run
            left -> right). 0 and 1 are corners shared by two edges and are
            rejected as ambiguous — use e.g. 0.01/0.99 or the adjacent side.
            When omitted, a free slot is chosen automatically with at least
            0.15 in between points (the first is the edge midpoint); a short,
            crowded edge errors instead of packing points together. Requires
            an explicit from_side.
        to_pos: Same for the to_side edge.
        waypoints: Optional page-coordinate [x, y] points (inches, origin
            bottom-left) the route must pass through. Supplying this creates
            a native 1-D route-controlled line instead of Visio's dynamic
            connector: explicit from_side/to_side are required, endpoints stay
            statically glued, and the points stay at their page positions
            when shapes move or the file is reopened. The first/last segments
            still leave/enter perpendicular to the pinned sides; between
            waypoints the route is orthogonal. No obstacle avoidance is done.
        label_pos: Optional 0..1 position along the routed path for the label
            (0 = begin, 1 = end). Works on dynamic and route-controlled
            connectors; the label stays at that absolute page point.
        label_offset_in: Optional perpendicular label offset in inches
            (positive = left of the direction of travel).
    """
    return await _worker.run(
        _client.connect_shapes, from_id, to_id, label, route,
        end_arrow, begin_arrow, page, line_pattern, line_weight_pt, line_color,
        from_side, to_side, from_pos, to_pos,
        waypoints, label_pos, label_offset_in,
    )


@mcp.tool()
@_tool_errors
async def set_connector_route(
    connector_id: int,
    waypoints: Optional[list[list[float]]] = None,
    label_pos: Optional[float] = None,
    label_offset_in: Optional[float] = None,
    page: Optional[str] = None,
) -> dict:
    """Change an existing connector's route and/or label position.

    With waypoints, a dynamic connector is replaced by a native 1-D
    route-controlled line preserving endpoints, sides, positions, arrows and
    line style; an existing manual route is rewritten in place. Without
    waypoints only the label moves. Pass [] to give a connector a default
    orthogonal route with no extra bends beyond the side stubs.

    Args:
        connector_id: The connector shape to adjust.
        waypoints: Page-coordinate [x, y] points (inches, bottom-left origin),
            at most 8. Requires a connector statically glued on both ends
            (from_side/to_side); dynamic-glue connectors are rejected.
        label_pos: Optional 0..1 position along the routed path for the label.
        label_offset_in: Optional perpendicular offset in inches (positive =
            left of travel).
    """
    return await _worker.run(
        _client.set_connector_route, connector_id, waypoints,
        label_pos, label_offset_in, page,
    )


@mcp.tool()
@_tool_errors
async def set_route_options(
    separate_lanes: bool = True,
    page: Optional[str] = None,
) -> dict:
    """Turn on native page-level connector lane separation.

    Sets the page layout cells LineAdjustFrom=1 (separate all lines) and
    LineAdjustTo=2 (never overlap lines). Visio's router then keeps different
    connectors out of the same lane, which removes short merge trunks and
    long co-linear overlaps without changing any connector's endpoints or
    glue. Pass separate_lanes=False to restore the page defaults. Use
    route_diagnostics to verify the result; manual routes via
    set_connector_route keep exact corridors where the heuristic is not
    enough.
    """
    return await _worker.run(_client.set_route_options, separate_lanes, page)


@mcp.tool()
@_tool_errors
async def pages(
    action: Literal["list", "add", "activate"],
    name: Optional[str] = None,
) -> dict:
    """List pages, add a page (optionally named), or activate a page by name."""
    return await _worker.run(_client.pages, action, name)


@mcp.tool()
@_tool_errors
async def auto_layout(
    style: Literal[
        "flowchart_tb", "flowchart_lr", "tree_tb", "tree_lr", "radial", "circular"
    ] = "flowchart_tb",
    spacing_in: float = 0.75,
    resize_page: bool = True,
    page: Optional[str] = None,
) -> dict:
    """Run Visio's automatic layout on the page — arranges shapes and
    re-routes connectors. Great after dropping shapes at rough positions.
    'flowchart_tb' = top-to-bottom flowchart, '_lr' = left-to-right."""
    return await _worker.run(_client.auto_layout, style, spacing_in, resize_page, page)


@mcp.tool()
@_tool_errors
async def set_page_size(
    width_in: Optional[float] = None,
    height_in: Optional[float] = None,
    orientation: Optional[Literal["portrait", "landscape"]] = None,
    fit_to_contents: bool = False,
    page: Optional[str] = None,
) -> dict:
    """Resize the page. Do this BEFORE building a wide architecture diagram —
    the default page is US Letter (8.5 x 11 in) and export_page_png crops to
    the page bounds, so shapes placed beyond them are invisible.

    Args:
        width_in / height_in: New page size in inches (e.g. 20 x 12 for a
            hybrid-cloud reference diagram).
        orientation: Print orientation; inferred from the size if omitted.
        fit_to_contents: True = shrink/grow the page to fit what's on it
            (alternative to explicit width/height).
    """
    return await _worker.run(
        _client.set_page_size, width_in, height_in, orientation, fit_to_contents, page
    )


@mcp.tool()
@_tool_errors
async def drop_text(
    text: str,
    x: float,
    y: float,
    width_in: Optional[float] = None,
    height_in: Optional[float] = None,
    font_size_pt: float = 10.0,
    bold: bool = False,
    text_color: Optional[str] = None,
    align: Literal["left", "center", "right"] = "center",
    page: Optional[str] = None,
) -> dict:
    """Drop a text-only label (no border, no fill) — for diagram titles,
    legends, and callouts. (x, y) is the CENTER in inches. Size is estimated
    from the text if omitted. Auto-layout ignores unconnected text shapes."""
    return await _worker.run(
        _client.drop_text, text, x, y, width_in, height_in,
        font_size_pt, bold, text_color, align, page,
    )


@mcp.tool()
@_tool_errors
async def add_container(
    label: str,
    member_ids: list[int],
    master: Optional[str] = None,
    padding_in: float = 0.4,
    page: Optional[str] = None,
) -> dict:
    """Wrap existing shapes in a real Visio container (for VNets, VPCs,
    subnets, resource groups, trust zones). The container is sized around its
    members plus padding, and members MOVE WITH the container afterwards.
    Nest zones by creating inner containers first, then outer ones.

    Args:
        label: Container heading text.
        member_ids: shape_ids to place inside.
        master: Optional container style name from Visio's built-in container
            stencil (defaults to the first available style).

    To recolor the container afterwards (e.g. brand colors per zone), use
    style_shape on the returned shape_id — theme guards are overridden.
    """
    return await _worker.run(
        _client.add_container, label, member_ids, master, padding_in, page
    )


@mcp.tool()
@_tool_errors
async def container_members(
    action: Literal["add", "remove"],
    container_id: int,
    member_ids: list[int],
    page: Optional[str] = None,
) -> dict:
    """Add shapes to or remove shapes from an existing container."""
    return await _worker.run(
        _client.container_members, action, container_id, member_ids, page
    )


@mcp.tool()
@_tool_errors
async def badge_container(
    container_id: int,
    icon: str,
    corner: Literal["top_left", "top_right", "bottom_left", "bottom_right"] = "top_left",
    size_in: float = 0.35,
    page: Optional[str] = None,
) -> dict:
    """Badge a container (zone) with a small header icon — e.g. the VNet icon
    on a VNet container, or the Subnet/Subscription icon on those boundaries.

    Placement is deterministic: the icon is dropped just inside the chosen
    corner of the container (computed from the container's bounds — no manual
    coordinate math) and added as a container member so it moves with the
    zone. `icon` resolves to a built-in Visio master first (e.g. 'Virtual
    Networks'), then to a labeled image in the local icon folder (e.g. a saved
    'subnet.svg' — see list_local_icons); errors with guidance when neither
    exists. Returns the badge's shape_id."""
    return await _worker.run(
        _client.badge_container, container_id, icon, corner, size_in, page
    )


@mcp.tool()
@_tool_errors
async def get_page_state(page: Optional[str] = None) -> dict:
    """List everything on a page: shape ids, master names, positions (inches,
    center), sizes, text, connector endpoints (from_id/to_id), and container
    membership (container_ids). Connectors also report how each end is glued:
    'dynamic' (Visio picks the side) or 'static' (a connection point), with
    the edge (from_side/to_side), the position along it (from_pos/to_pos,
    0..1) and the 1-based connection-point row (from_point/to_point), plus
    route_kind ('auto' dynamic connector or 'manual' route-controlled line).
    Call this to (re)discover shape ids, e.g. after open_document."""
    return await _worker.run(_client.get_page_state, page)


@mcp.tool()
@_tool_errors
async def route_diagnostics(page: Optional[str] = None) -> dict:
    """Read-only routing quality report for a page. No shapes are changed.

    For every connector it returns the actual routed polyline in page inches,
    route length, straight-line length and detour ratio, plus precise
    locations of:
      - crossings with other connectors (connector ids + page point),
      - collinear overlaps with other connectors (id + overlap length +
        segment),
      - node traversal / text occlusion (shape ids + inside length),
      - label conflicts (other connector or shape + length/area).
    Use it after connect_shapes / set_connector_route to verify that a lane
    change actually removed an overlap, or to find the connector ids that
    need a waypoint."""
    return await _worker.run(_client.route_diagnostics, page)


def main() -> None:
    try:
        mcp.run()  # stdio transport
    finally:
        _worker.shutdown(release=_client.release)


if __name__ == "__main__":
    main()
