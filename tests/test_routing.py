"""Pure routing geometry: orthogonal builder and route diagnostics."""

import math

import pytest

from visio_mcp import routing


def test_side_point_round_trip_on_all_sides():
    box = (1.0, 2.0, 5.0, 6.0)
    assert routing.side_point(box, "left", 0.25) == pytest.approx((1.0, 3.0))
    assert routing.side_point(box, "right", 0.75) == pytest.approx((5.0, 5.0))
    assert routing.side_point(box, "bottom", 0.5) == pytest.approx((3.0, 2.0))
    assert routing.side_point(box, "top", 0.0) == pytest.approx((1.0, 6.0))


def test_orthogonal_route_straight_vertical_cleans_collinear_points():
    route = routing.orthogonal_route((4.0, 8.0), (4.0, 4.0), "bottom", "top")
    assert route == pytest.approx([(4.0, 8.0), (4.0, 4.0)])


def test_orthogonal_route_leaves_and_enters_perpendicular():
    route = routing.orthogonal_route((4.0, 8.0), (7.0, 4.0), "bottom", "top")
    assert route[0] == (4.0, 8.0)
    assert route[-1] == (7.0, 4.0)
    # first segment leaves straight down, last segment enters straight down
    assert route[1][0] == pytest.approx(4.0)
    assert route[1][1] < route[0][1]
    assert route[-2][0] == pytest.approx(7.0)
    assert route[-2][1] > route[-1][1]
    for a, b in zip(route, route[1:]):
        assert abs(a[0] - b[0]) < 1e-9 or abs(a[1] - b[1]) < 1e-9, (a, b)


def test_orthogonal_route_passes_through_waypoints_in_order():
    waypoints = [(6.0, 6.0), (8.0, 3.0)]
    route = routing.orthogonal_route((4.0, 8.0), (9.0, 2.0), "bottom", "top",
                                     waypoints)
    for point in waypoints:
        assert point in route
    # order along the polyline is preserved
    assert route.index(waypoints[0]) < route.index(waypoints[1])


def test_orthogonal_route_loop_back_on_the_same_side():
    route = routing.orthogonal_route((5.0, 4.0), (5.0, 7.0), "right", "right",
                                     [(7.0, 5.5)])
    assert route[1][0] > 5.0  # leaves to the right
    assert route[-2][0] > 5.0  # enters from the right
    assert (7.0, 5.5) in route


def test_orthogonal_route_rejects_auto_or_bad_sides():
    with pytest.raises(ValueError, match="explicit"):
        routing.orthogonal_route((0, 0), (1, 1), "auto", "top")
    with pytest.raises(ValueError, match="explicit"):
        routing.orthogonal_route((0, 0), (1, 1), "bottom", "diagonal")


def test_point_at_fraction_walks_the_route():
    points = [(0.0, 0.0), (2.0, 0.0), (2.0, 2.0)]
    point, direction = routing.point_at_fraction(points, 0.5)
    assert point == pytest.approx((2.0, 0.0))
    assert direction == pytest.approx((1.0, 0.0))
    point, direction = routing.point_at_fraction(points, 1.0)
    assert point == pytest.approx((2.0, 2.0))
    assert direction == pytest.approx((0.0, 1.0))


def test_intersection_overlap_and_inside_rect_detectors():
    assert routing.segment_intersection(
        (0, 0), (2, 2), (0, 2), (2, 0)
    ) == pytest.approx((1.0, 1.0))
    assert routing.segment_intersection(
        (0, 0), (2, 0), (0, 1), (2, 1)
    ) is None
    assert routing.collinear_overlap_length(
        ((0, 0), (2, 0)), ((1, 0), (3, 0))
    ) == pytest.approx(1.0)
    assert routing.collinear_overlap_length(
        ((0, 0), (2, 0)), ((1, 0.5), (3, 0.5))
    ) == 0.0
    assert routing.segment_inside_rect(
        (0, 0.5), (2, 0.5), (0.5, 0, 1.5, 1)
    ) == pytest.approx(1.0)
    assert routing.segment_inside_rect(
        (0, 0), (2, 0), (0.5, 0, 1.5, 1), inset=0.02
    ) == 0.0


def test_diagnose_routes_reports_crossing_overlap_through_and_label():
    connectors = {
        1: {"from_id": 10, "to_id": 11, "points": [(0.0, 1.0), (4.0, 1.0)]},
        2: {"from_id": 12, "to_id": 13, "points": [(2.0, 0.0), (2.0, 2.0)]},
        3: {"from_id": 14, "to_id": 15, "points": [(1.0, 1.0), (3.0, 1.0)]},
    }
    nodes = {
        10: (0.0, 0.5, 0.5, 1.5),
        12: (1.5, 0.0, 2.5, 0.5),
        20: (3.5, 0.8, 4.5, 1.2),
        21: (1.6, 0.9, 2.4, 1.1),
    }
    labels = {1: (1.5, 0.8, 2.5, 1.2)}
    report = routing.diagnose_routes(connectors, nodes, labels)
    data = report["connectors"]

    assert data[1]["crossings"][0]["connector_id"] == 2
    assert data[2]["crossings"][0]["connector_id"] == 1
    assert data[1]["overlaps"][0]["connector_id"] == 3
    assert data[1]["overlaps"][0]["length_in"] == pytest.approx(2.0, abs=0.05)
    assert data[1]["through_shapes"][0]["shape_id"] == 20
    kinds = {conflict["kind"] for conflict in data[1]["label_conflicts"]}
    assert kinds == {"connector", "shape"}
    assert report["summary"]["flagged_connector_ids"]


def test_diagnose_routes_detour_ratio():
    connectors = {
        1: {"from_id": 1, "to_id": 2,
            "points": [(0.0, 0.0), (10.0, 0.0), (0.0, 0.0), (0.0, 10.0)]},
    }
    report = routing.diagnose_routes(connectors, {})
    ratio = report["connectors"][1]["detour_ratio"]
    assert ratio == pytest.approx(3.0)
    assert report["summary"]["detour_count"] == 1
