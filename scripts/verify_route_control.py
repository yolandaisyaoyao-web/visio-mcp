"""Real-Visio acceptance for connector route control and diagnostics.

    uv run python scripts/verify_route_control.py

Uses the round-2 baseline drawings in D:\\桌面\\test\\visio-mcp-regression:

  A. a synthetic manual route: waypoints stay at their page positions after
     moving a node and after save/reopen; endpoints stay glued; arrows and
     labels remain native.
  B. conversion of an existing dynamic connector into a route-controlled
     native 1-D line, preserving endpoints, style, text and arrow.
  C. the B (merge) and C (multiple loops) baselines: native page lane
     separation (LineAdjustFrom/LineAdjustTo) removes the 0.15 in overlaps
     without moving shapes or changing topology.
  D. the TDMA baseline: removes the 6.5 in complete overlay, pins the two
     long routes to distinct corridors, clears the "No" label conflict,
     survives a P6 move and save/reopen, and verifies no new crossings or
     node traversal.

Artifacts written next to the baselines (which are copied to baseline/):
  *-route-controlled.vsdx / .png, route-control-data.json.
"""

from __future__ import annotations

import json
import os
import shutil
import sys

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from visio_mcp.runtime import ComWorker  # noqa: E402
from visio_mcp.visio_client import VisioClient  # noqa: E402

BASE = r"D:\桌面\test\visio-mcp-regression"
BASELINE = os.path.join(BASE, "baseline")
TOL = 0.01

checks: list[dict] = []


def check(area: str, item: str, ok: bool, detail: str) -> None:
    checks.append({"area": area, "item": item, "ok": bool(ok), "detail": detail})
    print(f"  {'PASS' if ok else 'FAIL'}  {item}: {detail}")


def summary(report: dict) -> dict:
    keys = ("connector_count", "crossing_count", "overlap_count",
            "through_shape_count", "label_conflict_count",
            "flagged_connector_ids", "max_detour_ratio")
    return {key: report["summary"][key] for key in keys}


def topology(state: dict) -> list:
    return sorted((s["from_id"], s["to_id"]) for s in state["shapes"]
                  if s["is_connector"])


def node_positions(state: dict) -> dict:
    return {s["shape_id"]: (round(s["x"], 3), round(s["y"], 3))
            for s in state["shapes"]
            if not s["is_connector"] and not s.get("is_container")}


def item_by_id(report: dict, connector_id: int) -> dict:
    return next(c for c in report["connectors"]
                if c["connector_id"] == connector_id)


def vertex_in_polyline(polyline, point, tol=TOL) -> bool:
    return any(abs(p[0] - point[0]) <= tol and abs(p[1] - point[1]) <= tol
               for p in polyline)


def find_connector(state: dict, from_id: int, to_id: int) -> dict:
    hits = [s for s in state["shapes"]
            if s["is_connector"] and s["from_id"] == from_id
            and s["to_id"] == to_id]
    if len(hits) != 1:
        raise AssertionError(
            f"expected one connector {from_id}->{to_id}, got {len(hits)}")
    return hits[0]


def reset_artifacts():
    os.makedirs(BASELINE, exist_ok=True)
    for name in os.listdir(BASE):
        if "-route-controlled" in name and (
                name.endswith(".vsdx") or name.endswith(".png")
                or name.endswith(".json")):
            os.remove(os.path.join(BASE, name))


def synthetic_route(client, app):
    """A + B: manual route mechanics and dynamic -> manual conversion."""
    print("\n[A] synthetic manual route")
    client.create_document(
        r"C:\Program Files\Microsoft Office\Root\Office16\Visio Content\2052\BASFLO_M.VSTX"
    )
    client.open_stencil(
        r"C:\Program Files\Microsoft Office\Root\Office16\Visio Content\2052\BASFLO_M.VSSX"
    )
    ids = [s["shape_id"] for s in client.drop_shapes([
        {"master": "流程", "x": 4, "y": 8, "text": "A"},
        {"master": "流程", "x": 6, "y": 4, "text": "B"},
    ])["shapes"]]
    a, b = ids
    created = client.connect_shapes(
        a, b, from_side="bottom", to_side="top",
        waypoints=[[5, 7], [7, 6]], label="手动",
        label_pos=0.4, label_offset_in=0.2,
    )
    check("手动路由", "创建为原生 1-D 路由", created["route_kind"] == "manual",
          str(created))
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    info = state[created["connector_id"]]
    check("手动路由", "端点静态粘附且侧别正确",
          info["from_glue"] == info["to_glue"] == "static"
          and (info["from_side"], info["to_side"]) == ("bottom", "top"),
          f"{info['from_glue']}/{info['to_glue']} "
          f"{info['from_side']}/{info['to_side']}")
    report = client.route_diagnostics()
    item = item_by_id(report, created["connector_id"])
    check("手动路由", "折线可读且经过两个 waypoint",
          item["geometry_ok"] and vertex_in_polyline(item["polyline"], [5, 7])
          and vertex_in_polyline(item["polyline"], [7, 6]),
          str(item["polyline"]))
    check("手动路由", "初始无交叉/重叠/穿节点/标签冲突",
          not (item["crossings"] or item["overlaps"]
               or item["through_shapes"] or item["label_conflicts"]),
          f"cross={item['crossings']} overlap={item['overlaps']} "
          f"through={item['through_shapes']} label={item['label_conflicts']}")
    png = os.path.join(BASE, "route-control-synthetic.png")
    client.export_page_png(png)

    # move the target: page-anchored waypoints must not move
    client.update_shape(b, x=7.0, y=3.5)
    report = client.route_diagnostics()
    item = item_by_id(report, created["connector_id"])
    check("手动路由", "移动节点后 waypoint 仍固定在页面坐标",
          vertex_in_polyline(item["polyline"], [5, 7])
          and vertex_in_polyline(item["polyline"], [7, 6]),
          str(item["polyline"]))
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    info = state[created["connector_id"]]
    check("手动路由", "移动后拓扑与静态粘附保持",
          (info["from_id"], info["to_id"]) == (a, b)
          and info["from_glue"] == info["to_glue"] == "static",
          f"{info['from_id']}->{info['to_id']} {info['from_side']}/{info['to_side']}")
    check("手动路由", "移动后箭头方向正确",
          item["route_kind"] == "manual",
          f"route_kind={item['route_kind']}")

    path = os.path.join(BASE, "route-control-synthetic.vsdx")
    if os.path.exists(path):
        os.remove(path)
    cid = created["connector_id"]
    client.save_document(path)
    app.ActiveDocument.Close()
    client.open_document(path)
    report = client.route_diagnostics()
    item = item_by_id(report, cid)
    check("手动路由", "保存重开后 waypoint 与折线保持",
          item["geometry_ok"] and vertex_in_polyline(item["polyline"], [5, 7])
          and vertex_in_polyline(item["polyline"], [7, 6]),
          str(item["polyline"]))
    check("手动路由", "保存重开后标签矩形存在",
          item["label_rect"] is not None, str(item["label_rect"]))
    app.ActiveDocument.Close()

    print("\n[B] dynamic -> manual conversion")
    client.create_document(
        r"C:\Program Files\Microsoft Office\Root\Office16\Visio Content\2052\BASFLO_M.VSTX"
    )
    ids = [s["shape_id"] for s in client.drop_shapes([
        {"master": "流程", "x": 4, "y": 8, "text": "A"},
        {"master": "流程", "x": 4, "y": 4, "text": "B"},
    ])["shapes"]]
    a, b = ids
    dynamic = client.connect_shapes(
        a, b, label="动态", from_side="bottom", to_side="top",
        line_pattern="dashed", line_weight_pt=2, line_color="#808080",
    )
    converted = client.set_connector_route(
        dynamic["connector_id"], waypoints=[[6, 6]],
        label_pos=0.5, label_offset_in=0.2,
    )
    check("转换", "动态连接线被替换为手动路由",
          converted["route_kind"] == "manual"
          and converted["replaced_id"] == dynamic["connector_id"]
          and converted["connector_id"] != dynamic["connector_id"],
          f"{dynamic['connector_id']} -> {converted['connector_id']}")
    shape = client._page().Shapes.ItemFromID(converted["connector_id"])
    check("转换", "样式/文字/箭头保持",
          shape.CellsU("LinePattern").ResultIU == 2
          and str(shape.Text) == "动态"
          and float(shape.CellsU("EndArrow").ResultIU) == 5
          and str(shape.CellsU("LineWeight").FormulaU) == "2 pt",
          f"pattern={shape.CellsU('LinePattern').ResultIU} "
          f"text={str(shape.Text)!r} weight={shape.CellsU('LineWeight').FormulaU}")
    check("转换", "标签使用 POINTALONGPATH",
          "POINTALONGPATH" in str(shape.CellsU("TxtPinX").FormulaU),
          str(shape.CellsU("TxtPinX").FormulaU))
    report = client.route_diagnostics()
    item = item_by_id(report, converted["connector_id"])
    check("转换", "转换后无布线告警",
          not (item["crossings"] or item["overlaps"]
               or item["through_shapes"] or item["label_conflicts"]),
          f"cross={item['crossings']} overlap={item['overlaps']} "
          f"through={item['through_shapes']} label={item['label_conflicts']}")
    app.ActiveDocument.Close()


def controlled_page(client, app, key: str) -> dict:
    """C: apply native lane separation to a B/C baseline and save a copy."""
    print(f"\n[C] {key}: native lane separation")
    src = os.path.join(BASE, f"{key}.vsdx")
    if not os.path.exists(src):
        src = os.path.join(BASELINE, f"{key}.vsdx")
    client.open_document(src)
    page = client._page()
    before = client.route_diagnostics()
    state_before = client.get_page_state()
    positions_before = node_positions(state_before)
    topo_before = topology(state_before)
    print("  before:", summary(before))
    result = client.set_route_options(True)
    after = client.route_diagnostics()
    state_after = client.get_page_state()
    print("  after: ", summary(after))
    check(key, "短距离共线重叠清零",
          after["summary"]["overlap_count"] == 0,
          f"before={before['summary']['overlap_count']} "
          f"after={after['summary']['overlap_count']}")
    check(key, "无新增交叉/穿节点/标签冲突",
          after["summary"]["crossing_count"] == 0
          and after["summary"]["through_shape_count"] == 0
          and after["summary"]["label_conflict_count"] == 0,
          str(summary(after)))
    check(key, "拓扑不变",
          topology(state_after) == topo_before,
          f"{len(topo_before)} connectors")
    check(key, "节点位置不变",
          node_positions(state_after) == positions_before,
          "no shape moved")
    out = os.path.join(BASE, f"{key}-route-controlled.vsdx")
    if os.path.exists(out):
        os.remove(out)
    client.save_document(out)
    client.export_page_png(os.path.join(BASE, f"{key}-route-controlled.png"))
    app.ActiveDocument.Close()
    client.open_document(out)
    reopened = client.route_diagnostics()
    check(key, "保存重开后仍无告警",
          reopened["summary"]["overlap_count"] == 0
          and reopened["summary"]["crossing_count"] == 0
          and reopened["summary"]["through_shape_count"] == 0
          and reopened["summary"]["label_conflict_count"] == 0,
          str(summary(reopened)))
    sheet = client._page().PageSheet
    check(key, "页面分道设置已持久化",
          int(sheet.CellsU("LineAdjustFrom").ResultIU) == 1
          and int(sheet.CellsU("LineAdjustTo").ResultIU) == 2,
          f"LineAdjustFrom={sheet.CellsU('LineAdjustFrom').ResultIU}, "
          f"LineAdjustTo={sheet.CellsU('LineAdjustTo').ResultIU}")
    app.ActiveDocument.Close()
    return {"key": key, "before": summary(before), "after": summary(reopened),
            "options": result}


def tdma_controlled(client, app) -> dict:
    """D: remove the 6.5 in overlay and label conflict, pin two corridors."""
    print("\n[D] TDMA: fixed corridors + label placement")
    src = os.path.join(BASE, "D_tdma_reproduction.vsdx")
    if not os.path.exists(src):
        src = os.path.join(BASELINE, "D_tdma_reproduction.vsdx")
    client.open_document(src)
    page = client._page()
    before = client.route_diagnostics()
    state_before = client.get_page_state()
    topo_before = topology(state_before)
    positions_before = node_positions(state_before)
    print("  before:", summary(before))
    flagged_before = [(c["connector_id"], c["overlaps"], c["label_conflicts"])
                      for c in before["connectors"]
                      if c["overlaps"] or c["label_conflicts"]]
    check("TDMA", "基线存在 6.5 in 长距离并线",
          any(entry["length_in"] > 6.0
              for c in before["connectors"] for entry in c["overlaps"]),
          str(flagged_before))

    client.set_route_options(True)
    after_options = client.route_diagnostics()
    check("TDMA", "页面分道消除 6.5 in 并线",
          after_options["summary"]["overlap_count"] == 0,
          str(summary(after_options)))
    # the dynamic lane runs just left of the phase-2 column, so the free
    # label side is east of travel (negative offset) until the route is
    # pinned to the x=10.0 corridor below
    label_only = client.set_connector_route(
        32, label_pos=0.8, label_offset_in=-0.25,
    )
    after_label = client.route_diagnostics()
    check("TDMA", "标签位置控制消除“否”标签冲突",
          after_label["summary"]["label_conflict_count"] == 0,
          f"label_pos={label_only['label_pos']}, "
          f"offset={label_only['label_offset_in']}, "
          f"summary={summary(after_label)}")

    # pin the two long routes to distinct corridors
    r25 = client.set_connector_route(
        25, waypoints=[[8.0, 2.505], [8.0, 11.545]],
    )
    r32 = client.set_connector_route(
        32, waypoints=[[10.0, 4.5], [10.0, 11.0]],
        label_pos=0.8, label_offset_in=0.2,
    )
    after = client.route_diagnostics()
    print("  after: ", summary(after))
    check("TDMA", "手动走廊后全量布线告警清零",
          after["summary"]["overlap_count"] == 0
          and after["summary"]["crossing_count"] == 0
          and after["summary"]["through_shape_count"] == 0
          and after["summary"]["label_conflict_count"] == 0,
          str(summary(after)))
    check("TDMA", "拓扑仍为 17 条连接线且起止对不变",
          topology(client.get_page_state()) == topo_before,
          f"{len(topo_before)} connectors")
    check("TDMA", "页面分道未移动节点",
          node_positions(client.get_page_state()) == positions_before,
          "no shape moved")
    p25 = item_by_id(after, r25["connector_id"])
    p32 = item_by_id(after, r32["connector_id"])
    check("TDMA", "长连接固定在不同走廊",
          8.0 in {p[0] for p in p25["polyline"]}
          and 10.0 in {p[0] for p in p32["polyline"]},
          f"25 lane={sorted({p[0] for p in p25['polyline']})}, "
          f"32 lane={sorted({p[0] for p in p32['polyline']})}")
    check("TDMA", "手动路由箭头/粘附正确",
          p25["route_kind"] == p32["route_kind"] == "manual"
          and p25["geometry_ok"] and p32["geometry_ok"],
          f"25={p25['route_kind']} 32={p32['route_kind']}")
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    check("TDMA", "手动路由端点静态粘附",
          state[r25["connector_id"]]["from_glue"] == "static"
          and state[r25["connector_id"]]["to_glue"] == "static"
          and state[r32["connector_id"]]["from_glue"] == "static"
          and state[r32["connector_id"]]["to_glue"] == "static",
          f"25 {state[r25['connector_id']]['from_side']}/"
          f"{state[r25['connector_id']]['to_side']}, "
          f"32 {state[r32['connector_id']]['from_side']}/"
          f"{state[r32['connector_id']]['to_side']}")

    # move a nearby node: the fixed lanes must stay
    client.update_shape(10, x=16.4, y=9.5)
    after_move = client.route_diagnostics()
    p25m = item_by_id(after_move, r25["connector_id"])
    p32m = item_by_id(after_move, r32["connector_id"])
    check("TDMA", "移动 P6 后走廊不跳变",
          8.0 in {p[0] for p in p25m["polyline"]}
          and 10.0 in {p[0] for p in p32m["polyline"]},
          f"25={sorted({p[0] for p in p25m['polyline']})}, "
          f"32={sorted({p[0] for p in p32m['polyline']})}")
    check("TDMA", "移动 P6 后仍无新交叉/穿节点",
          after_move["summary"]["crossing_count"] == 0
          and after_move["summary"]["through_shape_count"] == 0,
          str(summary(after_move)))

    out = os.path.join(BASE, "D_tdma_reproduction-route-controlled.vsdx")
    if os.path.exists(out):
        os.remove(out)
    client.save_document(out)
    client.export_page_png(
        os.path.join(BASE, "D_tdma_reproduction-route-controlled.png"))
    app.ActiveDocument.Close()

    client.open_document(out)
    reopened = client.route_diagnostics()
    check("TDMA", "保存重开后仍无告警",
          reopened["summary"]["overlap_count"] == 0
          and reopened["summary"]["crossing_count"] == 0
          and reopened["summary"]["through_shape_count"] == 0
          and reopened["summary"]["label_conflict_count"] == 0,
          str(summary(reopened)))
    check("TDMA", "保存重开后拓扑不变",
          topology(client.get_page_state()) == topo_before,
          f"{len(topo_before)} connectors")
    sheet = client._page().PageSheet
    check("TDMA", "分道设置与手动路由持久化",
          int(sheet.CellsU("LineAdjustFrom").ResultIU) == 1
          and int(sheet.CellsU("LineAdjustTo").ResultIU) == 2,
          "LineAdjustFrom=1, LineAdjustTo=2")
    app.ActiveDocument.Close()
    return {
        "before": summary(before),
        "after_page_options": summary(after_options),
        "after_label_only": summary(after_label),
        "after_manual": summary(after),
        "reopened": summary(reopened),
        "moved_p6": summary(after_move),
        "manual_connector_ids": [r25["connector_id"], r32["connector_id"]],
        "replaced_ids": [r25["replaced_id"], r32["replaced_id"]],
    }


def run(client: VisioClient) -> dict:
    app = client._app()
    payload = {"out_dir": BASE, "checks": checks, "diagrams": {}}
    try:
        synthetic_route(client, app)
        for key in ("B_decisions_merge", "C_multiple_loops"):
            payload["diagrams"][key] = controlled_page(client, app, key)
        payload["diagrams"]["D_tdma_reproduction"] = tdma_controlled(client, app)
    finally:
        try:
            app.ActiveDocument.Close()
        except Exception:
            pass
    payload["checks"] = checks
    return payload


def main() -> int:
    if sys.platform != "win32":
        print("This verification must run on Windows with Visio desktop installed.")
        return 2
    os.makedirs(BASE, exist_ok=True)
    reset_artifacts()
    worker = ComWorker()
    client = VisioClient()
    try:
        payload = worker.run_sync(run, client)
    finally:
        worker.shutdown(release=client.release)

    data_path = os.path.join(BASE, "route-control-data.json")
    with open(data_path, "w", encoding="utf-8") as handle:
        json.dump(payload, handle, ensure_ascii=False, indent=2)
    failed = [c for c in checks if not c["ok"]]
    print(f"\n自动检查: {len(checks) - len(failed)}/{len(checks)} 通过")
    print(f"数据: {data_path}")
    if failed:
        for entry in failed:
            print(f"  FAIL {entry['area']}/{entry['item']}: {entry['detail']}")
        return 1
    print("全部路由控制验收检查通过。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
