"""Real-Visio verification for connect_shapes from_side / to_side.

    uv run python scripts/verify_connector_sides.py

Runs against a live Visio desktop (Windows). Builds a small algorithm
flowchart with side-pinned connectors plus one default (auto) connector and
checks that:

  - pinned endpoints statically glue to the connection point on the
    requested side (formula references Connections.X, ToPart >= 100),
  - default 'auto' endpoints stay dynamically glued (_WALKGLUE, ToPart 3),
  - several connectors on one edge get distinct auto-slotted positions at
    least EDGE_MIN_SLOT_GAP_IN apart (loop-back and the decision's No branch),
  - corner positions (0/1) are rejected without side effects, and explicit
    from_pos works,
  - endpoints stay on their pinned edge/position after a shape moves and
    after auto_layout,
  - get_page_state reports glue kind, edge, position and point row, matching
    the positions returned by connect_shapes,
  - an imported icon with no connection points gets a side point created,
  - an injected mid-connect failure rolls back both the connector and the
    connection point it created (real Visio undo state, not just a flag),
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

from visio_mcp import constants as C  # noqa: E402
from visio_mcp.errors import VisioMcpError  # noqa: E402
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

        def connect(src, dst, from_side, to_side, label=None,
                    from_pos=None, to_pos=None):
            return client.connect_shapes(
                src, dst, label=label, route="right_angle",
                from_side=from_side, to_side=to_side,
                from_pos=from_pos, to_pos=to_pos,
            )

        print("[1] connections: pinned flow + auto probe")
        conns = {
            "start": connect(start, init, "bottom", "top"),
            "init": connect(init, cond, "bottom", "top"),
            "yes": connect(cond, body, "bottom", "top", "Y"),
            "body": connect(body, incr, "bottom", "top"),
            # the loop-back occupies the decision's right midpoint first; the
            # No branch on the same edge must not share it
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

        def side_point(box, side, pos):
            """Page coordinate of `pos` along `side` of a bbox."""
            if side in ("left", "right"):
                x = box[0] if side == "left" else box[2]
                return (x, box[1] + (box[3] - box[1]) * pos)
            y = box[1] if side == "bottom" else box[3]
            return (box[0] + (box[2] - box[0]) * pos, y)

        incr_box = bbox(shape(incr))
        cond_box = bbox(shape(cond))
        end_box = bbox(shape(end))
        cond_height = float(shape(cond).CellsU("Height").ResultIU)
        # second slot = midpoint minus the minimum physical gap, as a fraction
        expected_no_pos = round(0.5 - C.EDGE_MIN_SLOT_GAP_IN / cond_height, 4)
        check(near(endpoint(conn["loop"], True), side_point(incr_box, "right", 0.5)),
              "loop: begin sits at i++'s right edge midpoint")
        check(near(endpoint(conn["loop"], False), side_point(cond_box, "right", 0.5)),
              "loop: end sits at the decision's right edge midpoint")
        check(near(endpoint(conn["no"], True),
                   side_point(cond_box, "right", expected_no_pos)),
              f"no: second connector auto-slotted at {expected_no_pos} "
              "(0.15 in from the midpoint)")
        check(near(endpoint(conn["no"], False), side_point(end_box, "left", 0.5)),
              "no: end sits at the end shape's left edge midpoint")
        loop_end = endpoint(conn["loop"], False)
        no_begin = endpoint(conn["no"], True)
        check(abs(loop_end[1] - no_begin[1]) >= C.EDGE_MIN_SLOT_GAP_IN - 1e-3,
              "two branches on the decision's right edge are >= 0.15 in apart")
        check(abs(conns["loop"]["to_pos"] - 0.5) < 1e-6
              and abs(conns["no"]["from_pos"] - expected_no_pos) < 1e-6,
              "connect_shapes returns the resolved edge positions")

        print("[2b] corner positions are rejected without side effects")
        before_shapes = int(page.Shapes.Count)
        before_points = int(shape(cond).RowCount(7))
        for corner_pos in (0, 1, 0.0005, 0.9995):
            try:
                client.connect_shapes(cond, end, from_side="right",
                                      from_pos=corner_pos)
                check(False, f"corner pos {corner_pos} must be rejected")
            except VisioMcpError as exc:
                check("corner" in str(exc),
                      f"corner pos {corner_pos} rejected with guidance")
        check(int(page.Shapes.Count) == before_shapes,
              "corner rejections dropped no connector")
        check(int(shape(cond).RowCount(7)) == before_points,
              "corner rejections created no connection point")

        print("[3] get_page_state reports glue kind, edge and point")
        state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
        loop_state = state[conns["loop"]["connector_id"]]
        no_state = state[conns["no"]["connector_id"]]
        check(loop_state["to_glue"] == "static" and loop_state["to_side"] == "right"
              and abs(loop_state["to_pos"] - 0.5) < 1e-6,
              "get_page_state: loop end -> static right edge @ 0.5")
        check(no_state["from_glue"] == "static" and no_state["from_side"] == "right"
              and abs(no_state["from_pos"] - expected_no_pos) < 1e-6,
              "get_page_state: No branch begin -> static right edge at the "
              "expected slot")
        check(abs(conns["no"]["from_pos"] - no_state["from_pos"]) < 1e-6,
              "returned edge position matches get_page_state exactly")
        check(no_state["from_point"] != loop_state["to_point"],
              "the two connectors use different connection-point rows")
        auto_state = state[conns["auto"]["connector_id"]]
        check(auto_state["from_glue"] == "dynamic"
              and auto_state["to_glue"] == "dynamic"
              and auto_state["from_side"] is None
              and auto_state["from_pos"] is None,
              "get_page_state: auto connector -> dynamic glue, no fixed edge")

        print("[4] move a node: pinned edge positions follow")
        shape(incr).CellsU("PinX").ResultIU = 6.0
        shape(incr).CellsU("PinY").ResultIU = 2.6
        incr_box = bbox(shape(incr))
        cond_box = bbox(shape(cond))
        check(near(endpoint(conn["loop"], True), side_point(incr_box, "right", 0.5)),
              "loop: begin still on i++'s right edge midpoint after the move")
        check(near(endpoint(conn["body"], False), side_point(incr_box, "top", 0.5)),
              "body: end still on i++'s top edge midpoint after the move")
        check(near(endpoint(conn["no"], True),
                   side_point(cond_box, "right", expected_no_pos)),
              "no: assigned slot kept on the decision's right edge after the move")
        before_auto = endpoint(conn["auto"], False)
        shape(auto_b).CellsU("PinX").ResultIU = 11.0
        shape(auto_b).CellsU("PinY").ResultIU = 8.0
        check(not near(before_auto, endpoint(conn["auto"], False)),
              "auto: dynamic end re-walked to a new side after the move")

        print("[5] auto_layout keeps pinned edge positions")
        client.auto_layout("flowchart_tb", 0.6, True)
        incr_box = bbox(shape(incr))
        cond_box = bbox(shape(cond))
        check(near(endpoint(conn["loop"], True), side_point(incr_box, "right", 0.5)),
              "loop: begin stays at i++'s right edge midpoint after auto_layout")
        check(near(endpoint(conn["loop"], False), side_point(cond_box, "right", 0.5))
              and near(endpoint(conn["no"], True),
                       side_point(cond_box, "right", expected_no_pos)),
              "both right-edge slots survive auto_layout")

        print("[6] imported icon: explicit pos + auto point creation")
        icon_path = os.path.join(tmp, "connector_sides_icon.svg")
        with open(icon_path, "w", encoding="utf-8") as fh:
            fh.write('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">'
                     '<rect width="100" height="100" fill="#4472C4"/></svg>')
        icon = client.import_image(icon_path, 10.5, 3.0, width_in=0.5)
        icon_shape = shape(icon["shape_id"])
        check(int(icon_shape.RowCount(7)) == 0,
              "imported icon starts with no connection points")
        icon_conn = client.connect_shapes(
            icon["shape_id"], auto_a, from_side="bottom", to_side="top",
            from_pos=0.25,
        )
        icon_connector = shape(icon_conn["connector_id"])
        check(int(icon_shape.RowCount(7)) == 1,
              "a side connection point was created on the icon")
        check(glued_to_connection_point(icon_connector, True),
              "icon connector is statically glued to the created point")
        check(end_part(icon_connector, True) == 100,
              "icon connector's begin ToPart == 100 (first connection point row)")
        icon_box = bbox(icon_shape)
        check(near(endpoint(icon_connector, True), side_point(icon_box, "bottom", 0.25)),
              "icon: explicit from_pos=0.25 lands 25% along the bottom edge")

        state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
        before_glue = {
            sid: (s.get("from_glue"), s.get("from_side"), s.get("from_pos"),
                  s.get("to_glue"), s.get("to_side"), s.get("to_pos"))
            for sid, s in state.items() if s["is_connector"]
        }

        print("[7] injected failure rolls back connector and connection point")
        probe = shape(icon["shape_id"])
        before_shapes = int(page.Shapes.Count)
        before_points = int(probe.RowCount(7))
        before_conns = sum(1 for s in state.values() if s["is_connector"])

        original_glue = client._glue_target

        def failing_glue(target_shape, target_side, target_pos):
            original_glue(target_shape, target_side, target_pos)  # creates the point
            raise RuntimeError("injected connector failure")

        client._glue_target = failing_glue
        try:
            try:
                client.connect_shapes(icon["shape_id"], auto_b, from_side="left")
                check(False, "injected failure must propagate")
            except Exception as exc:  # noqa: BLE001 — the injected RuntimeError
                check("injected connector failure" in str(exc),
                      "connect_shapes surfaced the injected failure")
        finally:
            del client._glue_target

        check(int(page.Shapes.Count) == before_shapes,
              "rollback removed the dropped connector shape")
        check(int(shape(icon["shape_id"]).RowCount(7)) == before_points,
              "rollback removed the connection point created before the failure")
        check(sum(1 for s in client.get_page_state()["shapes"] if s["is_connector"])
              == before_conns,
              "no orphan connector after rollback")

        print("[8] save, reopen, verify glue persisted")
        client.save_document(save_path)
        app.ActiveDocument.Close()
        reopened = client.open_document(save_path)
        created_docs.append(app.ActiveDocument)
        page = client._page()
        state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
        conn_ids = [sid for sid, s in state.items() if s["is_connector"]]
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
        after_glue = {
            sid: (s.get("from_glue"), s.get("from_side"), s.get("from_pos"),
                  s.get("to_glue"), s.get("to_side"), s.get("to_pos"))
            for sid, s in state.items() if s["is_connector"]
        }
        check(after_glue == before_glue,
              "reopened: glue kind, edge and edge position all persisted")

        print("[9] export PNG for a visual check")
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
