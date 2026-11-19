"""Real-Visio verification for connect_shapes from_side / to_side.

    uv run python scripts/verify_connector_sides.py

Runs against a live Visio desktop (Windows). Builds a small algorithm
flowchart with side-pinned connectors plus one default (auto) connector and
checks that:

  - pinned endpoints statically glue to the connection point on the
    requested side (formula references Connections.X, ToPart >= 100),
  - default 'auto' endpoints stay dynamically glued (_WALKGLUE, ToPart 3),
  - endpoints stay on their pinned sides after a shape moves — including a
    loop-back pinned right -> right — and after auto_layout,
  - an imported icon with no connection points gets a side point created,
  - glue survives save + reopen of the .vsdx,
  - the routed page still exports as a PNG (open it and eyeball the result).

The BASFLO template/stencil is taken from Visio's own content folder and
master names are resolved via Master.NameU, so it works with localized
(e.g. Chinese) Visio installs. All COM work runs on the ComWorker thread.
Only the documents this script created are closed; other open Visio
drawings are left untouched.
"""

from __future__ import annotations

import glob
import os
import sys
import tempfile

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

if hasattr(sys.stdout, "reconfigure"):  # localized master names on Windows
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

from visio_mcp.runtime import ComWorker  # noqa: E402
from visio_mcp.visio_client import VisioClient  # noqa: E402

TOL = 0.01  # inches — drawing/route rounding tolerance


def bbox(shape) -> tuple[float, float, float, float]:
    """(left, bottom, right, top) of a shape in page inches."""
    left = float(shape.CellsU("PinX").ResultIU) - float(shape.CellsU("LocPinX").ResultIU)
    bottom = float(shape.CellsU("PinY").ResultIU) - float(shape.CellsU("LocPinY").ResultIU)
    return (left, bottom,
            left + float(shape.CellsU("Width").ResultIU),
            bottom + float(shape.CellsU("Height").ResultIU))


def endpoint(conn, begin: bool) -> tuple[float, float]:
    axis = "Begin" if begin else "End"
    return (float(conn.CellsU(axis + "X").ResultIU),
            float(conn.CellsU(axis + "Y").ResultIU))


def glued_to_connection_point(conn, begin: bool) -> bool:
    axis = "Begin" if begin else "End"
    return "Connections." in str(conn.CellsU(axis + "X").FormulaU)


def glued_dynamically(conn, begin: bool) -> bool:
    axis = "Begin" if begin else "End"
    return "_WALKGLUE" in str(conn.CellsU(axis + "X").FormulaU)


def end_part(conn, begin: bool) -> int:
    """ToPart of the Begin/End Connect row (3 = PinX, >=100 = conn. point)."""
    for i in range(1, int(conn.Connects.Count) + 1):
        connect = conn.Connects.Item(i)
        if str(connect.FromCell.Name).startswith("Begin" if begin else "End"):
            return int(connect.ToPart)
    return -1


def near(a: tuple[float, float], b: tuple[float, float]) -> bool:
    return abs(a[0] - b[0]) <= TOL and abs(a[1] - b[1]) <= TOL


def find_basflo(app) -> tuple[str, str]:
    """Absolute (template, stencil) paths for the installed BASFLO family."""
    root = os.path.join(str(app.Path), "Visio Content")
    stencils = glob.glob(os.path.join(root, "*", "BASFLO_*.VSSX"))
    if not stencils:
        raise RuntimeError(f"No BASFLO stencil under {root!r}")
    stencil = sorted(stencils)[0]
    template = os.path.splitext(stencil)[0] + ".VSTX"
    return (template if os.path.exists(template) else ""), stencil


def master_names(app) -> dict[str, str]:
    """NameU -> localized Name for the open BASFLO stencil."""
    for i in range(1, int(app.Documents.Count) + 1):
        doc = app.Documents.Item(i)
        if int(doc.Type) == 2 and "BASFLO" in str(doc.Name).upper():
            return {str(m.NameU): str(m.Name) for m in doc.Masters}
    raise RuntimeError("BASFLO stencil not open")


def run_checks(client: VisioClient) -> dict:
    """All COM work — executed on the ComWorker's STA thread."""
    checks: list[tuple[bool, str]] = []

    def check(condition: bool, message: str) -> None:
        checks.append((bool(condition), message))

    app = client._app()
    template, stencil = find_basflo(app)
    created_docs = []
    created_stencils = []
    tmp = tempfile.gettempdir()
    png_path = os.path.join(tmp, "visio_mcp_connector_sides.png")
    save_path = os.path.join(tmp, "visio_mcp_connector_sides.vsdx")

    try:
        info = client.create_document(template or None)
        created_docs.append(app.ActiveDocument)
        checks.append((True, f"created document {info['document']!r}"))
        created_stencils.append(app.Documents.OpenEx(stencil, 6))
        names = master_names(app)
        print("    masters: " + ", ".join(
            f"{u}={names[u]}" for u in ("Start/End", "Process", "Decision") if u in names
        ))

        start_n, proc_n, dec_n = names["Start/End"], names["Process"], names["Decision"]
        dropped = client.drop_shapes([
            {"master": start_n, "x": 5.0, "y": 9.5, "text": "开始"},
            {"master": proc_n, "x": 5.0, "y": 8.2, "text": "i=0"},
            {"master": dec_n, "x": 5.0, "y": 6.8, "text": "i<n?"},
            {"master": proc_n, "x": 5.0, "y": 5.2, "text": "处理 a[i]"},
            {"master": proc_n, "x": 5.0, "y": 3.6, "text": "i++"},
            {"master": start_n, "x": 8.0, "y": 6.8, "text": "结束"},
            # isolated pair for the default 'auto' behaviour
            {"master": proc_n, "x": 10.5, "y": 8.0, "text": "auto-A"},
            {"master": proc_n, "x": 11.5, "y": 5.5, "text": "auto-B"},
        ])
        ids = [s["shape_id"] for s in dropped["shapes"]]
        start, init, cond, body, incr, end, auto_a, auto_b = ids
        page = client._page()

        def shape(shape_id):
            return page.Shapes.ItemFromID(int(shape_id))

        def connect(src, dst, from_side, to_side, label=None):
            return client.connect_shapes(
                src, dst, label, "right_angle", True, False, None,
                "solid", None, None, from_side, to_side,
            )

        print("[1] connections: pinned flow + auto probe")
        conns = {
            "start": connect(start, init, "bottom", "top"),
            "init": connect(init, cond, "bottom", "top"),
            "yes": connect(cond, body, "bottom", "top", "Y"),
            "body": connect(body, incr, "bottom", "top"),
            "loop": connect(incr, cond, "right", "right"),
            "no": connect(cond, end, "right", "left", "N"),
            "auto": client.connect_shapes(auto_a, auto_b),
        }
        conn = {key: shape(value["connector_id"]) for key, value in conns.items()}

        print("[2] glue kinds and pinned sides")
        for key in ("start", "init", "yes", "body", "loop", "no"):
            check(glued_to_connection_point(conn[key], True)
                  and glued_to_connection_point(conn[key], False),
                  f"{key}: both ends glued to side connection points")
            check(end_part(conn[key], True) >= 100 and end_part(conn[key], False) >= 100,
                  f"{key}: ToPart >= 100 for both ends")
        check(glued_dynamically(conn["auto"], True)
              and glued_dynamically(conn["auto"], False),
              "auto: both ends use dynamic _WALKGLUE glue")
        check(end_part(conn["auto"], True) == 3
              and end_part(conn["auto"], False) == 3,
              "auto: ToPart == 3 (PinX dynamic glue)")

        def mid_y(box):
            return (box[1] + box[3]) / 2

        def mid_x(box):
            return (box[0] + box[2]) / 2

        incr_box = bbox(shape(incr))
        cond_box = bbox(shape(cond))
        end_box = bbox(shape(end))
        check(near(endpoint(conn["loop"], True), (incr_box[2], mid_y(incr_box))),
              "loop: begin sits at i++'s right side midpoint")
        check(near(endpoint(conn["loop"], False), (cond_box[2], mid_y(cond_box))),
              "loop: end sits at the decision's right side midpoint")
        check(near(endpoint(conn["no"], True), (cond_box[2], mid_y(cond_box))),
              "no: begin sits at the decision's right side midpoint")
        check(near(endpoint(conn["no"], False), (end_box[0], mid_y(end_box))),
              "no: end sits at the end shape's left side midpoint")

        print("[3] move a node: pinned ends follow their side")
        shape(incr).CellsU("PinX").ResultIU = 6.0
        shape(incr).CellsU("PinY").ResultIU = 2.6
        incr_box = bbox(shape(incr))
        check(near(endpoint(conn["loop"], True), (incr_box[2], mid_y(incr_box))),
              "loop: begin still on i++'s right side after the move")
        check(near(endpoint(conn["body"], False), (mid_x(incr_box), incr_box[3])),
              "body: end still on i++'s top side after the move")
        before_auto = endpoint(conn["auto"], False)
        shape(auto_b).CellsU("PinX").ResultIU = 11.0
        shape(auto_b).CellsU("PinY").ResultIU = 8.0
        check(not near(before_auto, endpoint(conn["auto"], False)),
              "auto: dynamic end re-walked to a new side after the move")

        print("[4] auto_layout keeps pinned sides")
        client.auto_layout("flowchart_tb", 0.6, True)
        incr_box = bbox(shape(incr))
        check(near(endpoint(conn["loop"], True), (incr_box[2], mid_y(incr_box))),
              "loop: begin stays on i++'s right side after auto_layout")

        print("[5] imported icon (no connection points) gets a side point")
        icon_path = os.path.join(tmp, "connector_sides_icon.svg")
        with open(icon_path, "w", encoding="utf-8") as fh:
            fh.write('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">'
                     '<rect width="100" height="100" fill="#4472C4"/></svg>')
        icon = client.import_image(icon_path, 10.5, 3.0, width_in=0.5)
        icon_shape = shape(icon["shape_id"])
        check(int(icon_shape.RowCount(7)) == 0,
              "imported icon starts with no connection points")
        icon_conn = client.connect_shapes(
            icon["shape_id"], auto_a, None, "right_angle", True, False, None,
            "solid", None, None, "bottom", "top",
        )
        icon_connector = shape(icon_conn["connector_id"])
        check(int(icon_shape.RowCount(7)) == 1,
              "a side connection point was created on the icon")
        check(glued_to_connection_point(icon_connector, True),
              "icon connector is statically glued to the created point")
        check(end_part(icon_connector, True) == 100,
              "icon connector's begin ToPart == 100 (first connection point row)")

        print("[6] save, reopen, verify glue persisted")
        client.save_document(save_path)
        app.ActiveDocument.Close()
        reopened = client.open_document(save_path)
        created_docs.append(app.ActiveDocument)
        page = client._page()
        state = client.get_page_state()
        conn_ids = [s["shape_id"] for s in state["shapes"] if s["is_connector"]]
        check(len(conn_ids) == 8,
              f"{reopened['document']}: 8 connectors survived save/reopen")
        pinned = sum(
            1 for sid in conn_ids
            if glued_to_connection_point(page.Shapes.ItemFromID(sid), True)
            and glued_to_connection_point(page.Shapes.ItemFromID(sid), False)
        )
        check(pinned == 7, "reopened: 7 connectors still statically glued")
        dynamic = sum(
            1 for sid in conn_ids
            if glued_dynamically(page.Shapes.ItemFromID(sid), True)
            and glued_dynamically(page.Shapes.ItemFromID(sid), False)
        )
        check(dynamic == 1, "reopened: the auto connector is still dynamic")

        print("[7] export PNG for a visual check")
        result = client.export_page_png(png_path)
        check(os.path.getsize(result["path"]) > 1024,
              f"PNG exported to {result['path']} (open it to eyeball the routes)")
    finally:
        for doc in created_docs:
            try:
                doc.Close()
            except Exception:
                pass
        for doc in created_stencils:
            try:
                doc.Close()
            except Exception:
                pass
    return {"checks": checks}


def main() -> int:
    if sys.platform != "win32":
        print("This verification must run on Windows with Visio desktop installed.")
        return 2

    worker = ComWorker()
    client = VisioClient()
    try:
        result = worker.run_sync(run_checks, client)
    except Exception as exc:  # noqa: BLE001 — report, do not swallow
        import traceback

        traceback.print_exc()
        print(f"\nVerification aborted: {exc!r}")
        return 1
    finally:
        worker.shutdown(release=client.release)

    print()
    failed = [message for ok, message in result["checks"] if not ok]
    for ok, message in result["checks"]:
        print(f"  {'PASS' if ok else 'FAIL'}  {message}")
    print()
    if failed:
        print(f"{len(failed)} check(s) FAILED.")
        return 1
    print("All connector side checks passed on real Visio.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
