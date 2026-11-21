"""Connection-point edge detection across shape sizes.

The side of a connection point used to be decided with an absolute 0.01 in
tolerance. On shapes shorter than 1 in, any point within 1% of a corner was
reported as lying on the adjacent edge as well, and bottom/top wins that tie:
a genuine left-edge point near the bottom corner came back as "bottom". On
large shapes the same 0.01 in was far too tight to mean anything. These
regressions pin the matching to a fraction of each dimension instead.
"""

import pytest

from visio_mcp import constants as C
from visio_mcp.visio_client import VisioClient

from .fake_visio import FakeApplication

# (width, height) in inches: a large shape, the standard flowchart master,
# a small icon and a tiny imported icon.
SIZES = [
    (4.0, 2.0),
    (2.0, 1.0),
    (1.0, 0.75),
    (0.4, 0.3),
    (0.2, 0.1),
]
SIZE_IDS = ["4x2", "2x1", "1x0.75", "0.4x0.3", "0.2x0.1"]


@pytest.fixture
def app():
    return FakeApplication()


@pytest.fixture
def client(app):
    return VisioClient(app_factory=lambda: app)


@pytest.mark.parametrize("width,height", SIZES, ids=SIZE_IDS)
def test_left_point_near_bottom_corner_is_reported_left(width, height):
    # 1% along the left edge: a legal explicit from_pos (above the corner
    # epsilon), and on a large shape far more than 0.01 in above the corner.
    pos = 0.01
    side, got = VisioClient._side_of_point(0.0, height * pos, width, height)
    assert side == "left"
    assert got == pytest.approx(pos)


@pytest.mark.parametrize("width,height", SIZES, ids=SIZE_IDS)
def test_right_point_near_top_corner_is_reported_right(width, height):
    pos = 0.01
    side, got = VisioClient._side_of_point(
        width, height * (1.0 - pos), width, height
    )
    assert side == "right"
    assert got == pytest.approx(1.0 - pos)


@pytest.mark.parametrize("width,height", SIZES, ids=SIZE_IDS)
def test_edge_fraction_does_not_claim_the_adjacent_edge(width, height):
    # a bottom-edge point 0.2% in from the left corner is not on the left edge
    assert VisioClient._edge_fraction(
        "left", width * 0.002, 0.0, width, height
    ) is None
    # ... and a left-edge point 0.2% up from the bottom corner is not bottom
    assert VisioClient._edge_fraction(
        "bottom", 0.0, height * 0.002, width, height
    ) is None


@pytest.mark.parametrize("width,height", SIZES, ids=SIZE_IDS)
def test_edge_fraction_still_finds_points_exactly_on_an_edge(width, height):
    assert VisioClient._edge_fraction(
        "left", 0.0, height * 0.25, width, height
    ) == pytest.approx(0.25)
    assert VisioClient._edge_fraction(
        "bottom", width * 0.75, 0.0, width, height
    ) == pytest.approx(0.75)
    # float noise only, as returned by Visio after a recalc
    assert VisioClient._edge_fraction(
        "right", width + width * 1e-12, height * 0.5, width, height
    ) == pytest.approx(0.5)


def test_exact_corners_keep_the_documented_bottom_top_precedence():
    assert VisioClient._side_of_point(0.0, 0.0, 4.0, 2.0) == ("bottom", 0.0)
    assert VisioClient._side_of_point(4.0, 2.0, 4.0, 2.0) == ("top", 1.0)
    assert VisioClient._side_of_point(4.0, 0.0, 4.0, 2.0) == ("bottom", 1.0)


def test_edge_tolerance_relative_values():
    # 0.05% of the measured dimension, floor 1e-6 in for tiny shapes
    assert VisioClient._side_tolerance(4.0) == pytest.approx(4.0 * C.SIDE_POINT_TOLERANCE_FRAC)
    assert VisioClient._side_tolerance(0.001) == C.SIDE_POINT_TOLERANCE_MIN_IN
    assert C.SIDE_POINT_TOLERANCE_FRAC < C.EDGE_CORNER_EPS, (
        "a point created at the smallest accepted explicit position must not "
        "also match the neighbouring edge"
    )


def test_get_page_state_reports_near_corner_point_on_a_large_shape(client, app):
    """End-to-end: a left-edge point 0.2% up a 4x2 shape must not be
    misreported as a bottom-edge point (it is 0.004 in above the corner,
    inside the old absolute 0.01 in tolerance)."""
    client.create_document()
    client.open_stencil("BASFLO_U.VSSX")
    ids = [
        s["shape_id"]
        for s in client.drop_shapes([
            {"master": "Process", "x": 4, "y": 7},
            {"master": "Process", "x": 4, "y": 4},
        ])["shapes"]
    ]
    shape = app.ActivePage.shapes[0]
    shape.CellsU("Width").ResultIU = 4.0
    shape.CellsU("Height").ResultIU = 2.0

    result = client.connect_shapes(ids[0], ids[1], from_side="left", from_pos=0.002)
    assert result["from_pos"] == 0.002
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    endpoint = state[result["connector_id"]]
    assert endpoint["from_glue"] == "static"
    assert endpoint["from_side"] == "left"
    assert endpoint["from_pos"] == pytest.approx(0.002)


def test_get_page_state_reports_near_corner_point_on_default_master(client):
    """The same failure on an ordinary 1x0.75 in flowchart shape."""
    client.create_document()
    client.open_stencil("BASFLO_U.VSSX")
    ids = [
        s["shape_id"]
        for s in client.drop_shapes([
            {"master": "Process", "x": 4, "y": 7},
            {"master": "Process", "x": 4, "y": 4},
        ])["shapes"]
    ]
    result = client.connect_shapes(ids[0], ids[1], from_side="left", from_pos=0.005)
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    endpoint = state[result["connector_id"]]
    assert endpoint["from_side"] == "left"
    assert endpoint["from_pos"] == pytest.approx(0.005)
