"""Real-Visio regression: algorithm flowcharts, routing analysis, round-trip.

    uv run python scripts/regression_algorithm_flowcharts.py

Builds three standard test diagrams with side-pinned connectors

  A. basic sequential flow,
  B. multiple decisions and a two-way merge,
  C. multiple feedback loops (left-edge and right-edge loop-backs),

plus a simplified reproduction of the TDMA protocol flowchart, then:

  - checks topology (expected source/target pairs), glue kind, endpoint page
    positions against the pinned side/position, and arrow directions,
  - extracts each connector's routed polyline from its Geometry section and
    counts connector-connector crossings, collinear overlaps, segments that
    cut through other nodes (a text-occlusion risk) and route detour ratios,
  - moves nodes and re-checks that pinned endpoints follow their sides,
  - saves each drawing, reopens it and re-checks topology/glue/positions,
  - exports PNGs so the result can be inspected visually.

Artifacts (VSDX/PNG/JSON) are written outside the repository, to
D:\\桌面\\test\\visio-mcp-regression. Standard output is a PASS/FAIL summary.

Only geometry Visio exposes is used; no path-planning code lives here.
"""

from __future__ import annotations

import glob
import json
import math
import os
import sys
from datetime import datetime

if hasattr(sys.stdout, "reconfigure"):
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")

sys.path.insert(0, os.path.join(os.path.dirname(__file__), "..", "src"))

from visio_mcp import constants as C  # noqa: E402
from visio_mcp import routing  # noqa: E402
from visio_mcp.runtime import ComWorker  # noqa: E402
from visio_mcp.visio_client import VisioClient  # noqa: E402

OUT_DIR = r"D:\桌面\test\visio-mcp-regression"
TOL = 0.02          # inches — drawing rounding tolerance for page points
END_EPS = 0.05      # ignore intersections this close to a segment end
OVERLAP_MIN = 0.08  # inches — minimum collinear overlap worth reporting
INSIDE_MIN = 0.05   # inches — minimum segment length inside a node to report

# ---------------------------------------------------------------------------
# geometry helpers (pure Python, no COM)
# ---------------------------------------------------------------------------


def dist(a, b):
    return math.hypot(a[0] - b[0], a[1] - b[1])


def bbox(shape):
    """(left, bottom, right, top) in page inches."""
    px = float(shape.CellsU("PinX").ResultIU)
    py = float(shape.CellsU("PinY").ResultIU)
    lx = float(shape.CellsU("LocPinX").ResultIU)
    ly = float(shape.CellsU("LocPinY").ResultIU)
    w = float(shape.CellsU("Width").ResultIU)
    h = float(shape.CellsU("Height").ResultIU)
    return (px - lx, py - ly, px - lx + w, py - ly + h)


def side_point(box, side, pos):
    left, bottom, right, top = box
    if side == "left":
        return (left, bottom + (top - bottom) * pos)
    if side == "right":
        return (right, bottom + (top - bottom) * pos)
    if side == "bottom":
        return (left + (right - left) * pos, bottom)
    return (left + (right - left) * pos, top)


def on_boundary(point, box, tol=0.05):
    left, bottom, right, top = box
    x, y = point
    if not (left - tol <= x <= right + tol and bottom - tol <= y <= top + tol):
        return False
    return (abs(x - left) <= tol or abs(x - right) <= tol
            or abs(y - bottom) <= tol or abs(y - top) <= tol)


def rect_overlap_area(a, b):
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return max(0.0, w) * max(0.0, h)


# ---------------------------------------------------------------------------
# diagram specifications
# ---------------------------------------------------------------------------

# Node master aliases resolved to the localized BASFLO master names at runtime.
DIAGRAMS = [
    {
        "key": "A_basic_sequence",
        "title": "A 基本顺序流程",
        "page": (10, 12),
        "nodes": [
            ("S", "start", 5.0, 11.0, "开始"),
            ("P1", "process", 5.0, 9.6, "接收数据"),
            ("P2", "process", 5.0, 8.0, "校验数据"),
            ("P3", "process", 5.0, 6.4, "写入缓存"),
            ("E", "start", 5.0, 4.8, "结束"),
        ],
        # (from, to, from_side, to_side, from_pos, to_pos, label)
        "edges": [
            ("S", "P1", "bottom", "top", None, None, None),
            ("P1", "P2", "bottom", "top", None, None, None),
            ("P2", "P3", "bottom", "top", None, None, None),
            # explicit positions exercise custom slots on a plain chain
            ("P3", "E", "bottom", "top", 0.25, 0.75, None),
        ],
        "moves": [("P2", 1.0, 0.0)],
        "texts": [],
    },
    {
        "key": "B_decisions_merge",
        "title": "B 多重判断与汇合",
        "page": (14, 10),
        "nodes": [
            ("S", "start", 4.0, 9.0, "开始"),
            ("P0", "process", 4.0, 7.7, "初始化"),
            ("D1", "decision", 4.0, 6.3, "有效输入?"),
            ("A", "process", 4.0, 4.9, "处理 A"),
            ("D2", "decision", 8.5, 6.3, "需要复核?"),
            ("B", "process", 8.5, 4.9, "处理 B"),
            ("M", "process", 6.25, 3.4, "汇合处理"),
            ("E1", "start", 6.25, 2.0, "完成"),
            ("E2", "start", 12.3, 6.3, "拒绝结束"),
        ],
        "edges": [
            ("S", "P0", "bottom", "top", None, None, None),
            ("P0", "D1", "bottom", "top", None, None, None),
            ("D1", "A", "bottom", "top", None, None, "是"),
            ("D1", "D2", "right", "left", None, None, "否"),
            ("D2", "B", "bottom", "top", None, None, "是"),
            ("D2", "E2", "right", "left", None, None, "否"),
            # two inputs share the merge node's top edge -> auto slots
            ("A", "M", "bottom", "top", None, None, None),
            ("B", "M", "bottom", "top", None, None, None),
            ("M", "E1", "bottom", "top", None, None, None),
        ],
        "moves": [("B", 0.8, 0.0), ("E2", 0.3, 0.0)],
        "texts": [],
    },
    {
        "key": "C_multiple_loops",
        "title": "C 多循环反馈",
        "page": (14, 11),
        "nodes": [
            ("S", "start", 4.0, 9.8, "开始"),
            ("I1", "process", 4.0, 8.5, "初始化 i=0"),
            ("C1", "decision", 4.0, 7.1, "i<n?"),
            ("B1", "process", 4.0, 5.7, "处理 a[i]"),
            ("N1", "process", 4.0, 4.3, "i++"),
            ("C2", "decision", 8.5, 7.1, "校验通过?"),
            ("B2", "process", 8.5, 5.7, "二次处理"),
            ("N2", "process", 8.5, 4.3, "重试计数++"),
            ("E", "start", 12.3, 7.1, "结束"),
        ],
        "edges": [
            ("S", "I1", "bottom", "top", None, None, None),
            ("I1", "C1", "bottom", "top", None, None, None),
            ("C1", "B1", "bottom", "top", None, None, "是"),
            ("B1", "N1", "bottom", "top", None, None, None),
            # feedback 1: body retry, around the decision's right edge
            ("B1", "C1", "right", "right", None, None, "重试"),
            # feedback 2: counter loop, around the decision's left edge
            ("N1", "C1", "left", "left", None, None, None),
            # No branch shares C1's right edge with the retry connector
            ("C1", "C2", "right", "left", None, None, "否"),
            ("C2", "B2", "bottom", "top", None, None, "是"),
            ("B2", "N2", "bottom", "top", None, None, None),
            # feedback 3: second-stage counter loop, right edge
            ("N2", "C2", "right", "right", None, None, None),
            # No branch shares C2's right edge with the loop
            ("C2", "E", "right", "left", None, None, "否"),
        ],
        "moves": [("B1", 0.4, -0.2), ("C2", 0.6, -0.4)],
        "texts": [],
    },
    {
        "key": "D_tdma_reproduction",
        "title": "D TDMA 协议流程（简化复现）",
        "page": (22, 13),
        "nodes": [
            ("S", "start", 4.0, 11.6, "开始：接收到外部命令"),
            ("P1", "process", 4.0, 10.2, "Sink 广播 Beacon\n并监听 Join"),
            ("P2", "process", 4.0, 8.9, "Node 接收 Beacon"),
            ("D1", "decision", 4.0, 7.4, "有数据\n需要传输?"),
            ("P3", "process", 4.0, 5.9, "数据节点回复 Join"),
            ("D2", "decision", 4.0, 4.4, "Sink 计时器\n结束?"),
            ("P4", "process", 4.0, 2.8, "Sink 分配时隙\n并广播 Assign"),
            ("P5", "process", 11.0, 11.0, "数据节点\n按时隙传输数据"),
            ("D3", "decision", 11.0, 9.3, "有失败\n数据包?"),
            ("P6", "process", 15.5, 9.3, "失败包置于队顶\n等待重传"),
            ("P7", "process", 11.0, 7.7, "Sink 统计信息\n并回复 ACK"),
            ("P8", "process", 11.0, 6.1, "Node 接收 ACK"),
            ("D4", "decision", 11.0, 4.5, "满足\n终止条件?"),
            ("E", "start", 17.5, 4.5, "结束"),
        ],
        "edges": [
            ("S", "P1", "bottom", "top", None, None, None),
            ("P1", "P2", "bottom", "top", None, None, None),
            ("P2", "D1", "bottom", "top", None, None, None),
            ("D1", "P3", "bottom", "top", None, None, "是"),
            ("D1", "P1", "left", "left", None, None, "否"),
            ("P3", "D2", "bottom", "top", None, None, None),
            ("D2", "P4", "bottom", "top", None, None, "是"),
            ("D2", "D1", "right", "right", None, None, "否"),
            ("P4", "P5", "bottom", "top", None, None, None),
            ("P5", "D3", "bottom", "top", None, None, None),
            ("D3", "P7", "bottom", "top", None, None, "否"),
            ("D3", "P6", "right", "left", None, None, "是"),
            ("P6", "P5", "top", "right", None, None, "重传"),
            ("P7", "P8", "bottom", "top", None, None, None),
            ("P8", "D4", "bottom", "top", None, None, None),
            ("D4", "P5", "left", "left", None, None, "否"),
            ("D4", "E", "right", "left", None, None, "是"),
        ],
        "moves": [("P6", 0.9, -0.4), ("P3", 0.3, -0.2)],
        "texts": [("时隙预约阶段", 4.5, 12.3), ("数据包发送阶段", 13.5, 12.3)],
    },
]


# ---------------------------------------------------------------------------
# build / analyze
# ---------------------------------------------------------------------------


def iter_docs(app):
    for i in range(1, int(app.Documents.Count) + 1):
        yield app.Documents.Item(i)


def _doc_stem(doc):
    return os.path.splitext(str(doc.Name))[0].lower()


def _is_stencil(doc, stem):
    return int(doc.Type) == 2 and _doc_stem(doc) == stem.lower()


def localized_names(app, stencil):
    """NameU -> localized Name for the already-open BASFLO stencil."""
    stem = os.path.splitext(os.path.basename(stencil))[0]
    for doc in iter_docs(app):
        if _is_stencil(doc, stem):
            return {str(m.NameU): str(m.Name) for m in doc.Masters}
    raise RuntimeError(f"stencil {stem!r} is not open")


def open_stencil_doc(app, stencil):
    stem = os.path.splitext(os.path.basename(stencil))[0]
    for doc in iter_docs(app):
        if _is_stencil(doc, stem):
            return doc, False
    return app.Documents.OpenEx(stencil, C.STENCIL_OPEN_FLAGS), True


def find_basflo(app):
    root = os.path.join(str(app.Path), "Visio Content")
    stencils = glob.glob(os.path.join(root, "*", "BASFLO_*.VSSX"))
    if not stencils:
        raise RuntimeError(f"No BASFLO stencil under {root!r}")
    stencil = sorted(stencils)[0]
    template = os.path.splitext(stencil)[0] + ".VSTX"
    return (template if os.path.exists(template) else None), stencil


def add_check(checks, area, item, ok, detail, method="自动", quality=False):
    checks.append({"area": area, "item": item, "ok": bool(ok),
                   "detail": detail, "method": method, "quality": quality})


def analyze_stage(client, page, spec, built, stage):
    """All programmatic checks for one stage of a diagram."""
    checks = []
    metrics = {"stage": stage}
    app = client._app()
    state = client.get_page_state()
    infos = {s["shape_id"]: s for s in state["shapes"]}
    conn_infos = {sid: s for sid, s in infos.items() if s["is_connector"]}
    nodes = [s for s in state["shapes"] if not s["is_connector"]]

    # -- topology ------------------------------------------------------------
    expected_pairs = []
    for edge in spec["edges"]:
        expected_pairs.append((built["ids"][edge[0]], built["ids"][edge[1]]))
    actual_pairs = [(s["from_id"], s["to_id"]) for s in conn_infos.values()]
    topology_ok = sorted(actual_pairs, key=str) == sorted(expected_pairs, key=str)
    add_check(checks, "拓扑", "连接线数量与起止节点",
              len(conn_infos) == len(expected_pairs) and topology_ok,
              f"期望 {len(expected_pairs)} 条，实际 {len(conn_infos)} 条；"
              f"期望对={expected_pairs}，实际对={actual_pairs}")

    # map pair -> connector id (pairs are unique in these diagrams)
    by_pair = {}
    for cid, info in conn_infos.items():
        by_pair[(info["from_id"], info["to_id"])] = cid

    # -- glue, arrows, endpoints --------------------------------------------
    endpoint_failures = []
    glue_failures = []
    arrow_failures = []
    for edge in spec["edges"]:
        (src_key, dst_key, from_side, to_side,
         spec_from_pos, spec_to_pos, label) = edge
        src, dst = built["ids"][src_key], built["ids"][dst_key]
        cid = by_pair.get((src, dst))
        if cid is None:
            endpoint_failures.append(f"{src_key}->{dst_key}: 连接线缺失")
            continue
        info = conn_infos[cid]
        conn = page.Shapes.ItemFromID(cid)
        # glue kind
        if from_side != "auto" and info.get("from_glue") != "static":
            glue_failures.append(f"{src_key}->{dst_key}: begin 非静态粘附")
        if to_side != "auto" and info.get("to_glue") != "static":
            glue_failures.append(f"{src_key}->{dst_key}: end 非静态粘附")
        # reported side must match the requested side (near-corner regression)
        if from_side != "auto" and info.get("from_side") != from_side:
            glue_failures.append(
                f"{src_key}->{dst_key}: begin 侧={info.get('from_side')} 期望 {from_side}")
        if to_side != "auto" and info.get("to_side") != to_side:
            glue_failures.append(
                f"{src_key}->{dst_key}: end 侧={info.get('to_side')} 期望 {to_side}")
        for end_name, spec_pos, actual_pos in (
            ("begin", spec_from_pos, info.get("from_pos")),
            ("end", spec_to_pos, info.get("to_pos")),
        ):
            if spec_pos is not None and (
                actual_pos is None or abs(actual_pos - spec_pos) > 1e-6
            ):
                glue_failures.append(
                    f"{src_key}->{dst_key}: {end_name} 位置={actual_pos} "
                    f"期望 {spec_pos}")
        # arrow direction
        end_arrow = float(conn.CellsU("EndArrow").ResultIU)
        begin_arrow = float(conn.CellsU("BeginArrow").ResultIU)
        if not (abs(end_arrow - C.ARROW_FILLED_TRIANGLE) < 1e-6
                and abs(begin_arrow - C.ARROW_NONE) < 1e-6):
            arrow_failures.append(
                f"{src_key}->{dst_key}: EndArrow={end_arrow} BeginArrow={begin_arrow}")
        # page positions of the endpoints
        begin = (float(conn.CellsU("BeginX").ResultIU),
                 float(conn.CellsU("BeginY").ResultIU))
        end = (float(conn.CellsU("EndX").ResultIU),
               float(conn.CellsU("EndY").ResultIU))
        src_box = bbox(page.Shapes.ItemFromID(src))
        dst_box = bbox(page.Shapes.ItemFromID(dst))
        from_pos = info.get("from_pos")
        to_pos = info.get("to_pos")
        if from_pos is None:
            if not on_boundary(begin, src_box):
                endpoint_failures.append(
                    f"{src_key}->{dst_key}: begin {begin} 不在源图形边界")
        else:
            expected = side_point(src_box, from_side, from_pos)
            if dist(begin, expected) > TOL:
                endpoint_failures.append(
                    f"{src_key}->{dst_key}: begin {begin} != {expected} "
                    f"({from_side} @ {from_pos})")
        if to_pos is None:
            if not on_boundary(end, dst_box):
                endpoint_failures.append(
                    f"{src_key}->{dst_key}: end {end} 不在目标图形边界")
        else:
            expected = side_point(dst_box, to_side, to_pos)
            if dist(end, expected) > TOL:
                endpoint_failures.append(
                    f"{src_key}->{dst_key}: end {end} != {expected} "
                    f"({to_side} @ {to_pos})")
        if label and info.get("text") != label:
            glue_failures.append(
                f"{src_key}->{dst_key}: 标签 {info.get('text')!r} != {label!r}")

    add_check(checks, "拓扑", "静态粘附与端点侧别", not glue_failures,
              "; ".join(glue_failures)
              or f"{len(spec['edges'])} 条连接线的端点保持指定边与位置")
    add_check(checks, "拓扑", "箭头方向", not arrow_failures,
              "; ".join(arrow_failures)
              or f"{len(spec['edges'])} 条连接线 EndArrow=实心三角, BeginArrow=无")
    add_check(checks, "端点位置", "端点落在指定边的指定位置", not endpoint_failures,
              "; ".join(endpoint_failures)
              or f"{2 * len(spec['edges'])} 个端点误差 <= {TOL} in")

    # -- node overlap --------------------------------------------------------
    node_boxes = [(s["shape_id"], (s["x"] - s["width_in"] / 2,
                                   s["y"] - s["height_in"] / 2,
                                   s["x"] + s["width_in"] / 2,
                                   s["y"] + s["height_in"] / 2), s)
                  for s in nodes]
    overlaps = []
    for i in range(len(node_boxes)):
        for j in range(i + 1, len(node_boxes)):
            if node_boxes[i][2].get("is_text") or node_boxes[j][2].get("is_text"):
                continue
            area = rect_overlap_area(node_boxes[i][1], node_boxes[j][1])
            if area > 0.01:
                overlaps.append((node_boxes[i][0], node_boxes[j][0], round(area, 3)))
    add_check(checks, "重叠", "节点矩形互不重叠", not overlaps,
              f"重叠对: {overlaps}" if overlaps else "无节点重叠", quality=True)

    # -- polylines + quality diagnostics (shared with route_diagnostics) ------
    polylines = {}
    fallback = []
    connector_diag = {}
    for cid, info in conn_infos.items():
        points, ok = client._connector_polyline(page.Shapes.ItemFromID(cid))
        polylines[cid] = points
        if not ok:
            fallback.append(cid)
        connector_diag[cid] = {
            "from_id": info["from_id"],
            "to_id": info["to_id"],
            "points": points,
            "geometry_ok": ok,
            "route_kind": (
                "manual" if client._is_manual_connector(
                    page.Shapes.ItemFromID(cid)) else "auto"
            ),
        }
    metrics["geometry_fallback_connectors"] = fallback

    node_rects = {
        sid: box for sid, box, sinfo in node_boxes
        if not sinfo.get("is_container")
    }
    label_rects = {}
    for cid in polylines:
        rect = client._connector_label_rect(page.Shapes.ItemFromID(cid))
        if rect is not None:
            label_rects[cid] = rect
    report = routing.diagnose_routes(
        connector_diag, node_rects, label_rects,
        endpoint_eps=END_EPS, overlap_min=OVERLAP_MIN, inside_min=INSIDE_MIN,
        node_inset=C.ROUTE_NODE_INSET_IN,
        label_area_min=C.ROUTE_LABEL_AREA_MIN_IN2,
        detour_ratio=2.0,
    )
    cids = sorted(polylines)

    crossings = []
    seen_cross = set()
    for cid in cids:
        for entry in report["connectors"][cid]["crossings"]:
            key = tuple(sorted((cid, entry["connector_id"])))
            if key in seen_cross:
                continue
            seen_cross.add(key)
            crossings.append((key[0], key[1], entry["at"][0], entry["at"][1]))
    metrics["crossings"] = crossings
    add_check(checks, "交叉", "连接线之间无交叉", not crossings,
              f"交叉 {len(crossings)} 处: {crossings}" if crossings else "无交叉",
              quality=True)

    overlap_pairs = []
    seen_overlap = set()
    for cid in cids:
        for entry in report["connectors"][cid]["overlaps"]:
            key = tuple(sorted((cid, entry["connector_id"])))
            if key in seen_overlap:
                continue
            seen_overlap.add(key)
            overlap_pairs.append((key[0], key[1], entry["length_in"]))
    metrics["overlaps"] = overlap_pairs
    add_check(checks, "重叠", "连接线之间无共线重叠", not overlap_pairs,
              f"重叠 {overlap_pairs}" if overlap_pairs else "无共线重叠",
              quality=True)

    through = []
    for cid in cids:
        for entry in report["connectors"][cid]["through_shapes"]:
            through.append((cid, entry["shape_id"], entry["length_in"]))
    metrics["through_nodes"] = through
    add_check(checks, "文字遮挡", "连接线不穿越其他节点文字区", not through,
              f"穿越 {through}" if through else "无穿越节点文字区的线段",
              quality=True)

    label_issues = []
    for cid in cids:
        for entry in report["connectors"][cid]["label_conflicts"]:
            label_issues.append((
                cid,
                "连接线" if entry["kind"] == "connector" else "节点",
                entry["id"],
                entry.get("length_in", entry.get("area_in2")),
            ))
    metrics["label_issues"] = label_issues
    add_check(checks, "文字遮挡", "连接线标签不被节点/其他线遮挡", not label_issues,
              f"标签冲突 {label_issues}" if label_issues else "标签区域清晰",
              quality=True)

    route_stats = []
    detours = []
    for cid in cids:
        data = report["connectors"][cid]
        info = conn_infos[cid]
        ratio = data["detour_ratio"]
        route_stats.append((cid, info["from_id"], info["to_id"],
                            data["route_length"], ratio))
        if ratio is not None and ratio > 2.0:
            detours.append((cid, info["from_id"], info["to_id"],
                            data["route_length"], ratio))
    metrics["detours"] = detours
    metrics["route_stats"] = route_stats
    add_check(checks, "绕行", "路径长度不过度绕行 (<=2x 直线距离)", not detours,
              f"严重绕行 {detours}" if detours else
              f"无超过 2x 直线距离的绕行；最大 "
              f"{max((r[4] for r in route_stats if r[4] is not None), default=0):.2f}x",
              quality=True)

    metrics["connectors"] = {
        cid: {
            "from_id": info["from_id"],
            "to_id": info["to_id"],
            "from_side": info.get("from_side"),
            "to_side": info.get("to_side"),
            "from_pos": info.get("from_pos"),
            "to_pos": info.get("to_pos"),
            "label": info.get("text"),
            "label_rect": (
                [round(v, 3) for v in label_rect]
                if (label_rect := client._connector_label_rect(
                    page.Shapes.ItemFromID(cid))) else None
            ),
            "points": [(round(x, 3), round(y, 3)) for x, y in polylines[cid]],
        }
        for cid, info in conn_infos.items()
    }
    metrics["node_texts"] = {
        s["shape_id"]: s.get("text") for s in nodes if s.get("text")
    }
    metrics["page_size_actual"] = [
        round(float(page.PageSheet.CellsU("PageWidth").ResultIU), 3),
        round(float(page.PageSheet.CellsU("PageHeight").ResultIU), 3),
    ]
    metrics["checks"] = checks
    return metrics


def build_diagram(client, app, spec, names, template):
    """Create, connect, move, save, reopen and analyze one diagram."""
    key = spec["key"]
    vsdx = os.path.join(OUT_DIR, f"{key}.vsdx")
    png_initial = os.path.join(OUT_DIR, f"{key}-initial.png")
    png_final = os.path.join(OUT_DIR, f"{key}-final.png")
    for path in (vsdx, png_initial, png_final):
        if os.path.exists(path):
            os.remove(path)

    doc = client.create_document(template)
    client.set_page_size(spec["page"][0], spec["page"][1])
    nameu = {"start": "Start/End", "process": "Process", "decision": "Decision"}
    drops = [{"master": names[nameu[alias]], "x": x, "y": y, "text": text}
             for _, alias, x, y, text in spec["nodes"]]
    dropped = client.drop_shapes(drops)
    ids = {node[0]: info["shape_id"]
           for node, info in zip(spec["nodes"], dropped["shapes"])}
    text_ids = []
    for text, x, y in spec.get("texts", []):
        info = client.drop_text(text, x, y, font_size_pt=14, bold=True,
                                text_color="#1F4E79")
        text_ids.append(info["shape_id"])

    page = client._page()
    results = []
    for (src_key, dst_key, from_side, to_side,
         from_pos, to_pos, label) in spec["edges"]:
        kwargs = dict(label=label, route="right_angle",
                      from_side=from_side, to_side=to_side)
        if from_pos is not None:
            kwargs["from_pos"] = from_pos
        if to_pos is not None:
            kwargs["to_pos"] = to_pos
        result = client.connect_shapes(ids[src_key], ids[dst_key], **kwargs)
        results.append(result)

    built = {"ids": ids, "results": results, "document": doc}
    stages = {}
    stages["initial"] = analyze_stage(client, page, spec, built, "initial")
    client.export_page_png(png_initial)

    # -- move nodes ----------------------------------------------------------
    for node_key, dx, dy in spec["moves"]:
        shape = page.Shapes.ItemFromID(ids[node_key])
        x = float(shape.CellsU("PinX").ResultIU) + dx
        y = float(shape.CellsU("PinY").ResultIU) + dy
        client.update_shape(ids[node_key], x=x, y=y)
    stages["moved"] = analyze_stage(client, page, spec, built, "moved")

    # -- save / reopen -------------------------------------------------------
    saved = client.save_document(vsdx)
    app.ActiveDocument.Close()
    reopened = client.open_document(vsdx)
    page = client._page()
    stages["reopened"] = analyze_stage(client, page, spec, built, "reopened")
    client.export_page_png(png_final)
    app.ActiveDocument.Close()

    return {
        "key": key,
        "title": spec["title"],
        "page_size": spec["page"],
        "node_count": len(spec["nodes"]),
        "connector_count": len(spec["edges"]),
        "vsdx": saved["path"],
        "png_initial": png_initial,
        "png_final": png_final,
        "reopened": reopened["document"],
        "text_ids": text_ids,
        "results": [
            {k: v for k, v in r.items() if k in
             ("connector_id", "from_id", "to_id", "from_side", "to_side",
              "from_pos", "to_pos")}
            for r in results
        ],
        "stages": stages,
    }


def analyzer_selftest():
    """Validate the pure-Python geometry detectors before trusting them."""
    checks = []

    def check(item, ok, detail):
        checks.append({"area": "分析器自检", "item": item, "ok": bool(ok),
                       "detail": detail, "method": "自动", "quality": False})

    def report(**connectors):
        return routing.diagnose_routes(connectors, {})

    crossing = report(
        a={"from_id": 1, "to_id": 2, "points": [(0, 0), (2, 2)]},
        b={"from_id": 3, "to_id": 4, "points": [(0, 2), (2, 0)]},
    )
    check("两线相交检测", crossing["summary"]["crossing_count"] == 1,
          f"交叉数={crossing['summary']['crossing_count']}")
    parallel = report(
        a={"from_id": 1, "to_id": 2, "points": [(0, 0), (2, 0)]},
        b={"from_id": 3, "to_id": 4, "points": [(0, 1), (2, 1)]},
    )
    check("平行线不误报相交", parallel["summary"]["crossing_count"] == 0,
          f"交叉数={parallel['summary']['crossing_count']}")
    overlap = report(
        a={"from_id": 1, "to_id": 2, "points": [(0, 0), (2, 0)]},
        b={"from_id": 3, "to_id": 4, "points": [(1, 0), (3, 0)]},
    )
    check("共线重叠检测", overlap["summary"]["overlap_count"] == 1
          and overlap["connectors"]["a"]["overlaps"][0]["length_in"] == 1.0,
          str(overlap["connectors"]["a"]["overlaps"]))
    no_overlap = report(
        a={"from_id": 1, "to_id": 2, "points": [(0, 0), (2, 0)]},
        b={"from_id": 3, "to_id": 4, "points": [(0, 1), (2, 1)]},
    )
    check("平行不共线不误报重叠", no_overlap["summary"]["overlap_count"] == 0,
          f"重叠数={no_overlap['summary']['overlap_count']}")
    through = routing.diagnose_routes(
        {"a": {"from_id": 1, "to_id": 2, "points": [(0, 0.5), (2, 0.5)]}},
        {9: (0.5, 0, 1.5, 1)},
    )
    check("线段穿越节点检测", through["summary"]["through_shape_count"] == 1,
          str(through["connectors"]["a"]["through_shapes"]))
    hugging = routing.diagnose_routes(
        {"a": {"from_id": 1, "to_id": 2, "points": [(0, 0), (2, 0)]}},
        {9: (0.5, 0, 1.5, 1)},
    )
    check("贴边线段不计入穿越", hugging["summary"]["through_shape_count"] == 0,
          str(hugging["connectors"]["a"]["through_shapes"]))
    return checks


def near_corner_real_visio_check(client, app, names, template):
    """Real-Visio version of the near-corner side regression."""
    checks = []
    client.create_document(template)
    client.set_page_size(10, 8)
    ids = [s["shape_id"] for s in client.drop_shapes([
        {"master": names["Process"], "x": 4, "y": 6, "text": "大图形"},
        {"master": names["Process"], "x": 4, "y": 3, "text": "左目标"},
        {"master": names["Process"], "x": 8, "y": 3, "text": "右目标"},
    ])["shapes"]]
    page = client._page()
    shape = page.Shapes.ItemFromID(ids[0])
    shape.CellsU("Width").ResultIU = 4.0
    shape.CellsU("Height").ResultIU = 2.0
    left = client.connect_shapes(ids[0], ids[1], from_side="left", from_pos=0.002)
    right = client.connect_shapes(ids[0], ids[2], from_side="right", from_pos=0.998)
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    for key, result, side, pos, label in (
        ("左", left, "left", 0.002, "4x2 图形左边 0.002 处不被误判为底边"),
        ("右", right, "right", 0.998, "4x2 图形右边 0.998 处不被误判为顶边"),
    ):
        info = state[result["connector_id"]]
        add_check(checks, "近角点", label,
                  info.get("from_side") == side
                  and abs((info.get("from_pos") or -1) - pos) < 1e-6,
                  f"get_page_state -> side={info.get('from_side')} "
                  f"pos={info.get('from_pos')}（期望 {side} @ {pos}）")
    app.ActiveDocument.Close()

    client.create_document(template)
    ids = [s["shape_id"] for s in client.drop_shapes([
        {"master": names["Process"], "x": 4, "y": 6, "text": "默认尺寸"},
        {"master": names["Process"], "x": 4, "y": 3, "text": "目标"},
    ])["shapes"]]
    result = client.connect_shapes(ids[0], ids[1], from_side="left", from_pos=0.005)
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    info = state[result["connector_id"]]
    add_check(checks, "近角点", "默认 1x0.75 图形左边 0.005 处不被误判为底边",
              info.get("from_side") == "left"
              and abs((info.get("from_pos") or -1) - 0.005) < 1e-6,
              f"get_page_state -> side={info.get('from_side')} pos={info.get('from_pos')}")
    app.ActiveDocument.Close()

    # small imported icon (0.2 x 0.2 in): 0.002 along the left edge is 0.0004 in
    # above the corner — well inside the old absolute 0.01 in tolerance
    icon_path = os.path.join(OUT_DIR, "small-icon.svg")
    with open(icon_path, "w", encoding="utf-8") as fh:
        fh.write('<svg xmlns="http://www.w3.org/2000/svg" width="100" height="100">'
                 '<rect width="100" height="100" fill="#4472C4"/></svg>')
    client.create_document(template)
    icon = client.import_image(icon_path, 4, 6, width_in=0.2)
    proc = client.drop_shapes([
        {"master": names["Process"], "x": 4, "y": 3, "text": "目标"},
    ])["shapes"][0]
    result = client.connect_shapes(icon["shape_id"], proc["shape_id"],
                                   from_side="left", from_pos=0.002)
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    info = state[result["connector_id"]]
    add_check(checks, "近角点", "0.2x0.2 图标左边 0.002 处不被误判为底边",
              info.get("from_side") == "left"
              and abs((info.get("from_pos") or -1) - 0.002) < 1e-6,
              f"get_page_state -> side={info.get('from_side')} pos={info.get('from_pos')}")
    app.ActiveDocument.Close()
    return checks


def main():
    if sys.platform != "win32":
        print("This regression must run on Windows with Visio desktop installed.")
        return 2
    os.makedirs(OUT_DIR, exist_ok=True)

    worker = ComWorker()
    client = VisioClient()
    payload = {"generated_at": datetime.now().isoformat(timespec="seconds"),
               "out_dir": OUT_DIR, "diagrams": [], "near_corner": [],
               "analyzer_selftest": analyzer_selftest()}
    try:
        def run(_client):
            app = _client._app()
            template, stencil = find_basflo(app)
            stencil_doc, opened_here = open_stencil_doc(app, stencil)
            try:
                names = localized_names(app, stencil)
                payload["visio_version"] = str(app.Version)
                payload["visio_path"] = str(app.Path)
                payload["template"] = template
                payload["stencil"] = stencil
                payload["masters"] = names
                for spec in DIAGRAMS:
                    print(f"\n=== {spec['title']} ===")
                    diagram = build_diagram(_client, app, spec, names, template)
                    payload["diagrams"].append(diagram)
                    for stage, data in diagram["stages"].items():
                        failed = [c for c in data["checks"] if not c["ok"]]
                        print(f"  [{stage}] {len(data['checks']) - len(failed)}"
                              f"/{len(data['checks'])} 自动检查通过")
                        for check in failed:
                            print(f"    FAIL {check['area']}/{check['item']}: "
                                  f"{check['detail']}")
                print("\n=== 分析器自检 ===")
                for check in payload["analyzer_selftest"]:
                    print(f"  {'PASS' if check['ok'] else 'FAIL'} "
                          f"{check['item']}: {check['detail']}")
                print("\n=== 真实 Visio 近角点回归 ===")
                payload["near_corner"] = near_corner_real_visio_check(
                    _client, app, names, template)
                for check in payload["near_corner"]:
                    print(f"  {'PASS' if check['ok'] else 'FAIL'} {check['item']}: "
                          f"{check['detail']}")
            finally:
                if opened_here:
                    try:
                        stencil_doc.Close()
                    except Exception:
                        pass
        worker.run_sync(run, client)
    finally:
        worker.shutdown(release=client.release)

    data_path = os.path.join(OUT_DIR, "regression-data.json")
    with open(data_path, "w", encoding="utf-8") as fh:
        json.dump(payload, fh, ensure_ascii=False, indent=2)
    print(f"\n数据已写入 {data_path}")

    total = failed_functional = 0
    quality_total = quality_failed = 0
    for diagram in payload["diagrams"]:
        for stage_data in diagram["stages"].values():
            for check in stage_data["checks"]:
                total += 1
                if not check["ok"]:
                    if check.get("quality"):
                        quality_failed += 1
                    else:
                        failed_functional += 1
                if check.get("quality"):
                    quality_total += 1
    for check in payload["near_corner"]:
        total += 1
        failed_functional += 0 if check["ok"] else 1
    for check in payload["analyzer_selftest"]:
        total += 1
        failed_functional += 0 if check["ok"] else 1
    quality_passed = quality_total - quality_failed
    print(f"\n功能正确性检查: {total - quality_total - failed_functional}"
          f"/{total - quality_total} 通过")
    print(f"布线质量观察项: {quality_passed}/{quality_total} 无告警")
    if failed_functional:
        print(f"{failed_functional} 项功能检查未通过 —— 见上面明细。")
        return 1
    print("全部功能检查通过。请查看 PNG 做目视复核。")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
