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


def collinear_overlap_interval(seg_a: Sequence[Point], seg_b: Sequence[Point],
                               line_tol: float = 0.02) -> Optional[tuple[Point, Point]]:
    """Actual overlapping interval of two collinear segments, or None.

    The interval is expressed in page coordinates on segment ``a``'s line,
    from the lower to the higher projection along ``a``'s direction. It is
    the shared stretch of the two segments, not either segment as a whole.
    """
    (ax1, ay1), (ax2, ay2) = seg_a
    (bx1, by1), (bx2, by2) = seg_b
    adx, ady = ax2 - ax1, ay2 - ay1
    alen = math.hypot(adx, ady)
    if alen <= EPS:
        return None
    ux, uy = adx / alen, ady / alen
    bdx, bdy = bx2 - bx1, by2 - by1
    blen = math.hypot(bdx, bdy)
    if blen <= EPS:
        return None
    if abs(ux * (bdy / blen) - uy * (bdx / blen)) > 1e-9:
        return None
    cross = abs((bx1 - ax1) * uy - (by1 - ay1) * ux)
    if cross > line_tol:
        return None
    tb1 = (bx1 - ax1) * ux + (by1 - ay1) * uy
    tb2 = (bx2 - ax1) * ux + (by2 - ay1) * uy
    lo = max(0.0, min(tb1, tb2))
    hi = min(alen, max(tb1, tb2))
    if hi - lo <= EPS:
        return None
    return ((ax1 + ux * lo, ay1 + uy * lo), (ax1 + ux * hi, ay1 + uy * hi))


def collinear_overlap_length(seg_a: Sequence[Point], seg_b: Sequence[Point],
                             line_tol: float = 0.02) -> float:
    """Overlap length of two collinear segments (0 when not collinear)."""
    interval = collinear_overlap_interval(seg_a, seg_b, line_tol)
    return distance(*interval) if interval is not None else 0.0


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


def _continues_straight(a: Point, b: Point, c: Point) -> bool:
    """True when a -> b -> c is collinear AND keeps moving forward.

    A reversal (``c`` back towards ``a``) is deliberately False: such a leg
    is real geometry an out-and-back route depends on, and dropping the
    turning point would silently delete it.
    """
    abx, aby = b[0] - a[0], b[1] - a[1]
    bcx, bcy = c[0] - b[0], c[1] - b[1]
    scale = math.hypot(abx, aby) * math.hypot(bcx, bcy)
    if scale <= EPS:
        return False
    if abs(abx * bcy - aby * bcx) > EPS * scale:
        return False
    return (abx * bcx + aby * bcy) > EPS * scale


def _clean_path(points: Iterable[Point]) -> list[Point]:
    """Drop duplicate and redundant collinear intermediate points.

    A collinear point is only dropped when the path keeps moving in the same
    direction, i.e. when the straight line still covers it exactly. A
    reversal is kept: dropping its turning point would delete a deliberate
    out-and-back leg. Waypoints are not special-cased here; after cleaning,
    ``route_violations`` re-checks that every requested waypoint still lies
    on the route, in order.
    """
    path: list[Point] = []
    for point in points:
        if not path or distance(path[-1], point) > EPS:
            path.append(point)
    if len(path) < 3:
        return path

    out: list[Point] = [path[0]]
    for index in range(1, len(path) - 1):
        point = path[index]
        if _continues_straight(out[-1], point, path[index + 1]):
            continue
        out.append(point)
    out.append(path[-1])
    return out


def _unit(a: Point, b: Point) -> Point:
    dx, dy = b[0] - a[0], b[1] - a[1]
    length = math.hypot(dx, dy)
    return (dx / length, dy / length) if length > EPS else (0.0, 0.0)


def _position_along(points: Sequence[Point], point: Point) -> float:
    """Distance along the polyline of ``point``'s closest projection."""
    walked = 0.0
    best: Optional[tuple[float, float]] = None
    for a, b in zip(points, points[1:]):
        seg = distance(a, b)
        if seg <= EPS:
            continue
        ux, uy = (b[0] - a[0]) / seg, (b[1] - a[1]) / seg
        t = ((point[0] - a[0]) * ux + (point[1] - a[1]) * uy)
        t = max(0.0, min(seg, t))
        projected = (a[0] + ux * t, a[1] + uy * t)
        offset = distance(point, projected)
        if best is None or offset < best[0] - EPS:
            best = (offset, walked + t)
        walked += seg
    return best[1] if best is not None else 0.0


def _distance_to_polyline(point: Point, points: Sequence[Point]) -> float:
    if len(points) < 2:
        return distance(point, points[0]) if points else math.inf
    return min(point_segment_distance(point, seg)
               for seg in zip(points, points[1:]))


def route_violations(points: Sequence[Point], from_side: Optional[str] = None,
                     to_side: Optional[str] = None,
                     waypoints: Sequence[Point] = (),
                     *, tol: float = 1e-6) -> list[str]:
    """Human-readable problems with an orthogonal route, empty when valid.

    Checks that every segment is axis-aligned and non-degenerate, that the
    first segment leaves along ``from_side``'s outward normal and the last
    segment enters against ``to_side``'s normal, and that every waypoint is
    on the route in the given order.
    """
    path = list(points)
    if len(path) < 2:
        return ["the route needs at least two points"]
    problems: list[str] = []
    for index, (a, b) in enumerate(zip(path, path[1:])):
        dx, dy = abs(b[0] - a[0]), abs(b[1] - a[1])
        if dx <= tol and dy <= tol:
            problems.append(f"segment {index} has zero length")
        elif min(dx, dy) > tol:
            problems.append(f"segment {index} is not orthogonal")

    def _check_end(a: Point, b: Point, expected: Point, label: str) -> None:
        if distance(a, b) <= tol:
            return  # already reported as zero length
        ux, uy = _unit(a, b)
        if abs(ux - expected[0]) > tol or abs(uy - expected[1]) > tol:
            problems.append(
                f"the {label} segment does not leave along the pinned side"
            )

    if from_side in SIDE_NORMALS:
        _check_end(path[0], path[1], SIDE_NORMALS[from_side], "first")
    if to_side in SIDE_NORMALS:
        normal = SIDE_NORMALS[to_side]
        _check_end(path[-2], path[-1], (-normal[0], -normal[1]), "last")

    positions: list[float] = []
    for index, waypoint in enumerate(waypoints):
        if _distance_to_polyline(waypoint, path) > tol:
            problems.append(f"waypoint {index} {tuple(waypoint)} is not on the route")
        positions.append(_position_along(path, waypoint))
    for index in range(1, len(positions)):
        if positions[index] + tol < positions[index - 1]:
            problems.append("waypoints are out of order along the route")
            break
    return problems


def orthogonal_route(begin: Point, end: Point, from_side: str, to_side: str,
                     waypoints: Sequence[Point] = (),
                     stub_in: float = 0.25) -> list[Point]:
    """Build a right-angle polyline: begin -> stub -> waypoints -> stub -> end.

    The first/last segments leave/enter perpendicular to the pinned sides
    (``stub_in`` long).  Between anchors each pair is joined with one
    horizontal+vertical bend, alternating with the previous segment so the
    path does not zig-zag.  Waypoints are absolute page points; no obstacle
    avoidance or path search is performed here.

    Reversals are preserved and a forward-collinear intermediate point is
    only removed when the straight line still covers it. The finished path
    is validated (orthogonality, end directions, waypoint passage and order);
    a violation raises ValueError instead of writing a broken route.
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
    cleaned = _clean_path(path)
    problems = route_violations(
        cleaned,
        from_side if stub_in > 0 else None,
        to_side if stub_in > 0 else None,
        waypoints,
    )
    if problems:
        raise ValueError("could not build a valid orthogonal route: "
                         + "; ".join(problems))
    return cleaned


def _merge_overlap_intervals(
    intervals: Sequence[tuple[Point, Point]],
    gap_tol: float = 1e-6,
) -> list[tuple[Point, Point]]:
    """Merge intervals that lie on the same line into their shared extent.

    Several segment pairs of two connectors can report adjacent pieces of
    one physical overlap; diagnostics should state the real interval once.
    """
    groups: dict[tuple[float, float, float], dict] = {}
    for (x1, y1), (x2, y2) in intervals:
        dx, dy = x2 - x1, y2 - y1
        length = math.hypot(dx, dy)
        if length <= EPS:
            continue
        ux, uy = dx / length, dy / length
        # canonical direction so a reversed segment lands in the same group
        if ux < -EPS or (abs(ux) <= EPS and uy < 0):
            ux, uy = -ux, -uy
        offset = ux * y1 - uy * x1  # perpendicular constant: cross(u, p)
        key = (round(ux, 9), round(uy, 9), round(offset, 9))
        t1 = ux * x1 + uy * y1
        t2 = ux * x2 + uy * y2
        lo, hi = (t1, t2) if t1 <= t2 else (t2, t1)
        group = groups.setdefault(
            key, {"u": (ux, uy), "offset": offset, "ranges": []}
        )
        group["ranges"].append((lo, hi))

    merged: list[tuple[Point, Point]] = []
    for group in groups.values():
        ux, uy = group["u"]
        nx, ny = -uy, ux
        offset = group["offset"]
        ranges = sorted(group["ranges"])
        current_lo, current_hi = ranges[0]
        for lo, hi in ranges[1:]:
            if lo <= current_hi + gap_tol:
                current_hi = max(current_hi, hi)
            else:
                merged.append((
                    (current_lo * ux + offset * nx,
                     current_lo * uy + offset * ny),
                    (current_hi * ux + offset * nx,
                     current_hi * uy + offset * ny),
                ))
                current_lo, current_hi = lo, hi
        merged.append((
            (current_lo * ux + offset * nx, current_lo * uy + offset * ny),
            (current_hi * ux + offset * nx, current_hi * uy + offset * ny),
        ))
    return merged


def _non_orthogonal_segments(points: Sequence[Point],
                             tol: float) -> list[dict]:
    """Segments that are not axis-aligned, with their off-axis deviation."""
    violations: list[dict] = []
    for index, (a, b) in enumerate(zip(points, points[1:])):
        dx, dy = abs(b[0] - a[0]), abs(b[1] - a[1])
        if dx <= EPS and dy <= EPS:
            continue
        deviation = min(dx, dy)
        if deviation > tol:
            violations.append({
                "index": index,
                "from": [round(a[0], 3), round(a[1], 3)],
                "to": [round(b[0], 3), round(b[1], 3)],
                "deviation_in": round(deviation, 3),
            })
    return violations


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
    ortho_tol: float = 0.001,
) -> dict:
    """Quality metrics for routed connectors.

    ``connectors`` maps shape id -> ``{"from_id", "to_id", "points"}``,
    optionally with ``route_kind`` ('auto'/'manual') and ``geometry_ok``
    (False when the route fell back to a straight Begin-End line). Returns
    per-connector route length, detour ratio, every crossing point, every
    physical collinear-overlap interval, node traversal, label conflicts and
    — for manual routes — orthogonality plus whether the path needs
    readjusting after a node move. All in page coordinates.
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
        geometry_ok = bool(info.get("geometry_ok", True))
        route_kind = info.get("route_kind")
        orthogonal: Optional[bool] = None
        non_orthogonal: list[dict] = []
        if route_kind == "manual" and len(points) >= 2:
            non_orthogonal = _non_orthogonal_segments(points, ortho_tol)
            orthogonal = not non_orthogonal
        needs_readjust = route_kind == "manual" and orthogonal is False
        results[cid] = {
            "route_length": round(length, 3),
            "straight_length": round(straight, 3),
            "detour_ratio": round(length / straight, 3) if straight > EPS else None,
            "geometry_ok": geometry_ok,
            "route_kind": route_kind,
            "orthogonal": orthogonal,
            "needs_readjust": needs_readjust,
            "non_orthogonal_segments": non_orthogonal,
            "readjust_hint": (
                "The endpoints moved away from this page-anchored route; call "
                "set_connector_route(connector_id, waypoints=[...]) to restore "
                "an orthogonal path."
            ) if needs_readjust else None,
            "crossings": [],
            "overlaps": [],
            "through_shapes": [],
            "label_conflicts": [],
        }

    for i, ca in enumerate(valid_ids):
        points_a = connectors[ca]["points"]
        ends_a = (points_a[0], points_a[-1])
        for cb in valid_ids[i + 1:]:
            points_b = connectors[cb]["points"]
            ends_b = (points_b[0], points_b[-1])

            # Every distinct crossing point between the pair. A contact is
            # only ignored when it sits at a glued end of BOTH connectors
            # (a shared endpoint on a shape); contacts at an interior bend
            # are real crossings and are reported.
            found: list[Point] = []
            for sa in zip(points_a, points_a[1:]):
                if seg_length(sa) < EPS:
                    continue
                for sb in zip(points_b, points_b[1:]):
                    if seg_length(sb) < EPS:
                        continue
                    point = segment_intersection(*sa, *sb)
                    if point is None:
                        continue
                    if (any(distance(point, end) <= endpoint_eps for end in ends_a)
                            and any(distance(point, end) <= endpoint_eps
                                    for end in ends_b)):
                        continue
                    if any(distance(point, other) <= 1e-6 for other in found):
                        continue  # the same bend hit through two segments
                    found.append(point)
            found.sort(key=lambda point: _position_along(points_a, point))
            for point in found:
                at = [round(point[0], 3), round(point[1], 3)]
                results[ca]["crossings"].append({"connector_id": cb, "at": at})
                results[cb]["crossings"].append({"connector_id": ca,
                                                 "at": list(at)})

            # Every physical overlap interval, merged per line.
            intervals: list[tuple[Point, Point]] = []
            for sa in zip(points_a, points_a[1:]):
                if seg_length(sa) < EPS:
                    continue
                for sb in zip(points_b, points_b[1:]):
                    if seg_length(sb) < EPS:
                        continue
                    interval = collinear_overlap_interval(sa, sb)
                    if interval is not None:
                        intervals.append(interval)
            merged = [interval for interval in _merge_overlap_intervals(intervals)
                      if distance(*interval) > overlap_min]
            # order each interval along A's travel direction
            ordered: list[tuple[Point, Point]] = []
            for start, end in merged:
                if (_position_along(points_a, end)
                        < _position_along(points_a, start)):
                    start, end = end, start
                ordered.append((start, end))
            ordered.sort(key=lambda interval: _position_along(points_a,
                                                              interval[0]))
            for start, end in ordered:
                length = distance(start, end)
                segment = [[round(start[0], 3), round(start[1], 3)],
                           [round(end[0], 3), round(end[1], 3)]]
                results[ca]["overlaps"].append({
                    "connector_id": cb,
                    "length_in": round(length, 3),
                    "interval": [list(segment[0]), list(segment[1])],
                    "segment": [list(segment[0]), list(segment[1])],
                })
                results[cb]["overlaps"].append({
                    "connector_id": ca,
                    "length_in": round(length, 3),
                    "interval": [list(segment[0]), list(segment[1])],
                    "segment": [list(segment[0]), list(segment[1])],
                })

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
        or data["label_conflicts"] or data["needs_readjust"]
        or (data["detour_ratio"] is not None and data["detour_ratio"] > detour_ratio)
    ]
    geometry_failed = [cid for cid in ids if not results[cid]["geometry_ok"]]
    needs_readjust = [cid for cid in ids if results[cid]["needs_readjust"]]
    crossing_entries = sum(len(data["crossings"]) for data in results.values())
    overlap_entries = sum(len(data["overlaps"]) for data in results.values())
    summary = {
        "connector_count": len(ids),
        "crossing_count": crossing_entries // 2,
        "overlap_count": sum(1 for data in results.values()
                             if data["overlaps"]) // 2,
        "overlap_interval_count": overlap_entries // 2,
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
        "detour_ratio_threshold": detour_ratio,
        "geometry_failed_ids": geometry_failed,
        "diagnostics_complete": not geometry_failed,
        "manual_needs_readjust_ids": needs_readjust,
        "manual_needs_readjust_count": len(needs_readjust),
        "orthogonal_tolerance_in": ortho_tol,
        "note": ((
            "Geometry extraction failed for connector(s) "
            f"{geometry_failed}; their diagnostics use a straight Begin-End "
            "fallback and are incomplete (diagnostics_complete=false)."
        ) if geometry_failed else None),
    }
    return {"connectors": results, "summary": summary}


def filter_report(report: dict, connector_id: int) -> dict:
    """Restrict a diagnose_routes report to one connector.

    The per-connector data still names the other connectors it crosses or
    overlaps; the summary counts are scoped to the selected connector while
    completeness markers keep referring to the whole page.
    """
    data = report["connectors"][connector_id]
    summary = dict(report["summary"])
    threshold = summary.get("detour_ratio_threshold", 2.0)
    summary.update({
        "connector_count": 1,
        "crossing_count": len(data["crossings"]),
        "overlap_count": len({entry["connector_id"]
                              for entry in data["overlaps"]}),
        "overlap_interval_count": len(data["overlaps"]),
        "through_shape_count": len(data["through_shapes"]),
        "label_conflict_count": len(data["label_conflicts"]),
        "detour_count": (1 if data["detour_ratio"] is not None
                         and data["detour_ratio"] > threshold else 0),
        "flagged_connector_ids": (
            [connector_id] if connector_id in summary["flagged_connector_ids"]
            else []),
        "max_detour_ratio": data["detour_ratio"],
        "manual_needs_readjust_ids": (
            [connector_id] if data["needs_readjust"] else []),
        "manual_needs_readjust_count": 1 if data["needs_readjust"] else 0,
        "filtered_connector_id": connector_id,
    })
    return {"connectors": {connector_id: data}, "summary": summary}
