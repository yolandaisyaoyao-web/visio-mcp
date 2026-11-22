"""Pure geometry for connector routing and route diagnostics.

No COM here: every point is ``(x, y)`` in page inches (origin bottom-left,
y upward), every rectangle is ``(left, bottom, right, top)``.  The COM side
lives in :mod:`visio_mcp.visio_client`; this module is unit-testable on any
platform and is shared with the real-Visio regression scripts.
"""

from __future__ import annotations

import math
from typing import Iterable, Optional, Sequence

Point = tuple[float, float]
Rect = tuple[float, float, float, float]

EPS = 1e-9

SIDE_NORMALS: dict[str, Point] = {
    "top": (0.0, 1.0),
    "bottom": (0.0, -1.0),
    "left": (-1.0, 0.0),
    "right": (1.0, 0.0),
}


def distance(a: Point, b: Point) -> float:
    return math.hypot(a[0] - b[0], a[1] - b[1])


def polyline_length(points: Sequence[Point]) -> float:
    return sum(distance(a, b) for a, b in zip(points, points[1:]))


def side_point(rect: Rect, side: str, pos: float) -> Point:
    """Point at fraction ``pos`` along one side of a page-space rectangle."""
    left, bottom, right, top = rect
    if side == "left":
        return (left, bottom + (top - bottom) * pos)
    if side == "right":
        return (right, bottom + (top - bottom) * pos)
    if side == "bottom":
        return (left + (right - left) * pos, bottom)
    if side == "top":
        return (left + (right - left) * pos, top)
    raise ValueError(f"unknown side {side!r}")


def point_at_fraction(points: Sequence[Point], fraction: float) -> tuple[Point, Point]:
    """Point at ``fraction`` of the polyline's length, plus its direction.

    ``fraction`` is clamped to 0..1.  Returns the point and the unit
    direction of the segment containing it.
    """
    if not points:
        raise ValueError("empty polyline")
    if len(points) == 1:
        return points[0], (1.0, 0.0)
    fraction = max(0.0, min(1.0, float(fraction)))
    total = polyline_length(points)
    if total <= EPS:
        return points[0], (1.0, 0.0)
    wanted = total * fraction
    walked = 0.0
    for a, b in zip(points, points[1:]):
        seg = distance(a, b)
        if seg <= EPS:
            continue
        if walked + seg >= wanted:
            t = (wanted - walked) / seg
            point = (a[0] + (b[0] - a[0]) * t, a[1] + (b[1] - a[1]) * t)
            return point, ((b[0] - a[0]) / seg, (b[1] - a[1]) / seg)
        walked += seg
    a, b = points[-2], points[-1]
    seg = distance(a, b) or 1.0
    return b, ((b[0] - a[0]) / seg, (b[1] - a[1]) / seg)


def seg_length(seg: Sequence[Point]) -> float:
    return distance(seg[0], seg[1])


def segment_intersection(p1: Point, p2: Point, p3: Point, p4: Point) -> Optional[Point]:
    """Proper intersection point of two segments, or None."""
    x1, y1 = p1
    x2, y2 = p2
    x3, y3 = p3
    x4, y4 = p4
    denom = (x1 - x2) * (y3 - y4) - (y1 - y2) * (x3 - x4)
    if abs(denom) < 1e-12:
        return None
    t = ((x1 - x3) * (y3 - y4) - (y1 - y3) * (x3 - x4)) / denom
    u = ((x1 - x3) * (y1 - y2) - (y1 - y3) * (x1 - x2)) / denom
    if -1e-9 <= t <= 1 + 1e-9 and -1e-9 <= u <= 1 + 1e-9:
        return (x1 + t * (x2 - x1), y1 + t * (y2 - y1))
    return None


def point_segment_distance(point: Point, seg: Sequence[Point]) -> float:
    ax, ay = seg[0]
    bx, by = seg[1]
    dx, dy = bx - ax, by - ay
    length2 = dx * dx + dy * dy
    if length2 <= 1e-12:
        return distance(point, seg[0])
    t = max(0.0, min(1.0, ((point[0] - ax) * dx + (point[1] - ay) * dy) / length2))
    return distance(point, (ax + t * dx, ay + t * dy))


def collinear_overlap_length(seg_a: Sequence[Point], seg_b: Sequence[Point],
                             line_tol: float = 0.02) -> float:
    """Overlap length of two collinear segments (0 when not collinear)."""
    (ax1, ay1), (ax2, ay2) = seg_a
    (bx1, by1), (bx2, by2) = seg_b
    adx, ady = ax2 - ax1, ay2 - ay1
    alen = math.hypot(adx, ady)
    if alen <= EPS:
        return 0.0
    ux, uy = adx / alen, ady / alen
    bdx, bdy = bx2 - bx1, by2 - by1
    blen = math.hypot(bdx, bdy)
    if blen <= EPS:
        return 0.0
    if abs(ux * bdy - uy * bdx) > 1e-9:
        return 0.0
    cross = abs((bx1 - ax1) * uy - (by1 - ay1) * ux)
    if cross > line_tol:
        return 0.0
    tb1 = (bx1 - ax1) * ux + (by1 - ay1) * uy
    tb2 = (bx2 - ax1) * ux + (by2 - ay1) * uy
    lo = max(0.0, min(tb1, tb2))
    hi = min(alen, max(tb1, tb2))
    return max(0.0, hi - lo)


def segment_inside_rect(p: Point, q: Point, rect: Rect, inset: float = 0.0) -> float:
    """Length of the segment strictly inside an inset rectangle (Liang-Barsky)."""
    left, bottom, right, top = rect
    left += inset
    bottom += inset
    right -= inset
    top -= inset
    if right <= left or top <= bottom:
        return 0.0
    x0, y0 = p
    dx, dy = q[0] - x0, q[1] - y0
    t0, t1 = 0.0, 1.0
    for edge_p, edge_q in ((-dx, x0 - left), (dx, right - x0),
                           (-dy, y0 - bottom), (dy, top - y0)):
        if abs(edge_p) < 1e-12:
            if edge_q < 0:
                return 0.0
        else:
            t = edge_q / edge_p
            if edge_p < 0:
                t0 = max(t0, t)
            else:
                t1 = min(t1, t)
            if t0 > t1:
                return 0.0
    return max(0.0, t1 - t0) * math.hypot(dx, dy)


def rect_overlap_area(a: Rect, b: Rect) -> float:
    w = min(a[2], b[2]) - max(a[0], b[0])
    h = min(a[3], b[3]) - max(a[1], b[1])
    return max(0.0, w) * max(0.0, h)


def _dir(a: Point, b: Point) -> Optional[str]:
    if abs(a[1] - b[1]) <= EPS and abs(a[0] - b[0]) > EPS:
        return "h"
    if abs(a[0] - b[0]) <= EPS and abs(a[1] - b[1]) > EPS:
        return "v"
    return None


def _clean_path(points: Iterable[Point]) -> list[Point]:
    """Drop duplicate and collinear intermediate points."""
    path: list[Point] = []
    for point in points:
        if not path or distance(path[-1], point) > EPS:
            path.append(point)
    out: list[Point] = []
    for point in path:
        if len(out) >= 2:
            a, b = out[-2], out[-1]
            if ((abs(a[0] - b[0]) <= EPS and abs(b[0] - point[0]) <= EPS)
                    or (abs(a[1] - b[1]) <= EPS and abs(b[1] - point[1]) <= EPS)):
                out[-1] = point
                continue
        out.append(point)
    return out


def orthogonal_route(begin: Point, end: Point, from_side: str, to_side: str,
                     waypoints: Sequence[Point] = (),
                     stub_in: float = 0.25) -> list[Point]:
    """Build a right-angle polyline: begin -> stub -> waypoints -> stub -> end.

    The first/last segments leave/enter perpendicular to the pinned sides
    (``stub_in`` long).  Between anchors each pair is joined with one
    horizontal+vertical bend, alternating with the previous segment so the
    path does not zig-zag.  Waypoints are absolute page points; no obstacle
    avoidance or path search is performed here.
    """
    if from_side not in SIDE_NORMALS or to_side not in SIDE_NORMALS:
        raise ValueError("manual routes need explicit top/bottom/left/right sides")
    if stub_in < 0:
        raise ValueError("stub_in must be >= 0")
    f = SIDE_NORMALS[from_side]
    t = SIDE_NORMALS[to_side]
    anchors: list[Point] = [begin]
    if stub_in > 0:
        anchors.append((begin[0] + f[0] * stub_in, begin[1] + f[1] * stub_in))
    anchors.extend(waypoints)
    if stub_in > 0:
        anchors.append((end[0] + t[0] * stub_in, end[1] + t[1] * stub_in))
    anchors.append(end)

    path: list[Point] = [anchors[0]]
    for target in anchors[1:]:
        current = path[-1]
        if distance(current, target) <= EPS:
            continue
        if _dir(current, target):
            path.append(target)
            continue
        previous = _dir(path[-2], current) if len(path) >= 2 else None
        if previous == "v":
            path.append((target[0], current[1]))
            path.append(target)
        elif previous == "h":
            path.append((current[0], target[1]))
            path.append(target)
        elif abs(f[1]) > EPS:  # leaving vertically: turn horizontally first
            path.append((target[0], current[1]))
            path.append(target)
        else:
            path.append((current[0], target[1]))
            path.append(target)
    return _clean_path(path)


def diagnose_routes(
    connectors: dict[int, dict],
    node_rects: dict[int, Rect],
    label_rects: Optional[dict[int, Rect]] = None,
    *,
    endpoint_eps: float = 0.05,
    overlap_min: float = 0.08,
    inside_min: float = 0.05,
    node_inset: float = 0.02,
    label_area_min: float = 0.02,
    detour_ratio: float = 2.0,
) -> dict[int, dict]:
    """Quality metrics for routed connectors.

    ``connectors`` maps shape id -> ``{"from_id", "to_id", "points"}``.
    Returns per-connector route length, detour ratio, crossings, collinear
    overlaps, node traversal and label conflicts, all in page coordinates.
    """
    label_rects = label_rects or {}
    results: dict[int, dict] = {}
    ids = sorted(connectors)
    valid_ids = [cid for cid in ids if len(connectors[cid].get("points") or []) >= 2]

    for cid in ids:
        info = connectors[cid]
        points = info.get("points") or []
        length = polyline_length(points)
        straight = distance(points[0], points[-1]) if len(points) >= 2 else 0.0
        results[cid] = {
            "route_length": round(length, 3),
            "straight_length": round(straight, 3),
            "detour_ratio": round(length / straight, 3) if straight > EPS else None,
            "crossings": [],
            "overlaps": [],
            "through_shapes": [],
            "label_conflicts": [],
        }

    for i, ca in enumerate(valid_ids):
        points_a = connectors[ca]["points"]
        for cb in valid_ids[i + 1:]:
            points_b = connectors[cb]["points"]
            crossing = None
            for sa in zip(points_a, points_a[1:]):
                if seg_length(sa) < EPS:
                    continue
                for sb in zip(points_b, points_b[1:]):
                    if seg_length(sb) < EPS:
                        continue
                    point = segment_intersection(*sa, *sb)
                    if point is None:
                        continue
                    # ignore contacts at either segment's endpoint (shared
                    # glue targets), not every point lying on a segment
                    if min(distance(point, sa[0]), distance(point, sa[1]),
                           distance(point, sb[0]), distance(point, sb[1])) <= endpoint_eps:
                        continue
                    crossing = [round(point[0], 3), round(point[1], 3)]
                    break
                if crossing:
                    break
            if crossing:
                results[ca]["crossings"].append({"connector_id": cb, "at": crossing})
                results[cb]["crossings"].append({"connector_id": ca, "at": crossing})

            overlap = 0.0
            overlap_seg = None
            for sa in zip(points_a, points_a[1:]):
                for sb in zip(points_b, points_b[1:]):
                    length = collinear_overlap_length(sa, sb)
                    if length > overlap:
                        overlap = length
                        overlap_seg = [list(sa[0]), list(sa[1])]
            if overlap > overlap_min:
                entry_a = {"connector_id": cb, "length_in": round(overlap, 3),
                           "segment": overlap_seg}
                entry_b = {"connector_id": ca, "length_in": round(overlap, 3),
                           "segment": overlap_seg}
                results[ca]["overlaps"].append(entry_a)
                results[cb]["overlaps"].append(entry_b)

    for cid in valid_ids:
        points = connectors[cid]["points"]
        info = connectors[cid]
        endpoints = {info.get("from_id"), info.get("to_id")}
        for sid, rect in node_rects.items():
            if sid in endpoints:
                continue
            hit = 0.0
            for seg in zip(points, points[1:]):
                hit = max(hit, segment_inside_rect(*seg, rect, inset=node_inset))
            if hit > inside_min:
                results[cid]["through_shapes"].append(
                    {"shape_id": sid, "length_in": round(hit, 3)})

        label = label_rects.get(cid)
        if label:
            for other in valid_ids:
                if other == cid:
                    continue
                hit = 0.0
                for seg in zip(connectors[other]["points"],
                               connectors[other]["points"][1:]):
                    hit = max(hit, segment_inside_rect(*seg, label))
                if hit > inside_min:
                    results[cid]["label_conflicts"].append(
                        {"kind": "connector", "id": other,
                         "length_in": round(hit, 3)})
            for sid, rect in node_rects.items():
                area = rect_overlap_area(label, rect)
                if area > label_area_min:
                    results[cid]["label_conflicts"].append(
                        {"kind": "shape", "id": sid, "area_in2": round(area, 3)})

    flagged = [
        cid for cid, data in results.items()
        if data["crossings"] or data["overlaps"] or data["through_shapes"]
        or data["label_conflicts"]
        or (data["detour_ratio"] is not None and data["detour_ratio"] > detour_ratio)
    ]
    summary = {
        "connector_count": len(ids),
        "crossing_count": sum(len(data["crossings"]) for data in results.values()) // 2,
        "overlap_count": sum(len(data["overlaps"]) for data in results.values()) // 2,
        "through_shape_count": sum(len(data["through_shapes"])
                                   for data in results.values()),
        "label_conflict_count": sum(len(data["label_conflicts"])
                                    for data in results.values()),
        "detour_count": sum(1 for data in results.values()
                            if data["detour_ratio"] is not None
                            and data["detour_ratio"] > detour_ratio),
        "flagged_connector_ids": flagged,
        "max_detour_ratio": max(
            (data["detour_ratio"] for data in results.values()
             if data["detour_ratio"] is not None), default=None),
    }
    return {"connectors": results, "summary": summary}
