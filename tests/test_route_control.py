"""Route-controlled connectors, label placement and route diagnostics."""

import math

import pytest

from visio_mcp.errors import VisioMcpError
from visio_mcp.visio_client import VisioClient

from .fake_visio import FakeApplication


@pytest.fixture
def app():
    return FakeApplication()


@pytest.fixture
def client(app):
    return VisioClient(app_factory=lambda: app)


def _two_shapes(client, first=(4, 8), second=(6, 4)):
    client.create_document()
    client.open_stencil("BASFLO_U.VSSX")
    return [
        s["shape_id"]
        for s in client.drop_shapes([
            {"master": "Process", "x": first[0], "y": first[1], "text": "A"},
            {"master": "Process", "x": second[0], "y": second[1], "text": "B"},
        ])["shapes"]
    ]


def _move_glued_end(shape, x, y):
    """Move a 1-D line's End like Visio does when the glued node moves.

    The fake does not run Visio's glue solver, so the shape frame (Pin,
    LocPin, Width, Angle) is recomputed here; the page-anchored geometry
    formulas then keep their interior vertices while the last vertex follows
    the new End.
    """
    begin_x = float(shape.CellsU("BeginX").ResultIU)
    begin_y = float(shape.CellsU("BeginY").ResultIU)
    shape.CellsU("EndX").ResultIU = x
    shape.CellsU("EndY").ResultIU = y
    length = math.hypot(x - begin_x, y - begin_y)
    shape.CellsU("Width").ResultIU = length
    shape.CellsU("PinX").ResultIU = (begin_x + x) / 2.0
    shape.CellsU("PinY").ResultIU = (begin_y + y) / 2.0
    shape.CellsU("LocPinX").ResultIU = length / 2.0
    shape.CellsU("LocPinY").ResultIU = 0.0
    shape.CellsU("Angle").ResultIU = math.atan2(y - begin_y, x - begin_x)


def test_manual_route_creates_native_1d_line_with_vertices(client, app):
    a, b = _two_shapes(client)
    result = client.connect_shapes(
        a, b, from_side="bottom", to_side="top", waypoints=[[6, 6]],
        label="R", label_pos=0.25, label_offset_in=0.1,
    )
    assert result["route_kind"] == "manual"
    assert result["from_pos"] == 0.5 and result["to_pos"] == 0.5
    shape = app.ActivePage.Shapes.ItemFromID(result["connector_id"])
    assert shape.OneD
    assert app.ActivePage.glue_log == [
        (shape.ID, "BeginX", a, "Connections.X3"),
        (shape.ID, "EndX", b, "Connections.X4"),
    ]
    assert [row.row_type for row in shape.geom_rows] == [137, 138, 139, 139, 139]
    # interior vertices are page-anchored formulas; last row is the glued end
    assert shape.geom_rows[2].x.FormulaU.startswith("((")
    assert "BeginX" in shape.geom_rows[2].x.FormulaU
    assert shape.geom_rows[-1].x.FormulaU == "Width"
    assert shape.CellsU("User.VisioMcpRoute").ResultIU == 1
    # native arrows/style/text
    assert shape.CellsU("EndArrow").ResultIU == 5
    assert shape.CellsU("BeginArrow").ResultIU == 0
    assert shape.CellsU("LinePattern").ResultIU == 1
    assert str(shape.Text) == "R"
    # label follows the routed path and is counter-rotated
    assert "POINTALONGPATH" in shape.CellsU("TxtPinX").FormulaU
    assert shape.CellsU("TxtAngle").FormulaU == "-Angle"
    assert "TEXTWIDTH" in shape.CellsU("TxtWidth").FormulaU

    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    info = state[result["connector_id"]]
    assert info["is_connector"] is True
    assert info["route_kind"] == "manual"
    assert (info["from_id"], info["to_id"]) == (a, b)
    assert info["from_glue"] == info["to_glue"] == "static"
    assert (info["from_side"], info["to_side"]) == ("bottom", "top")


def test_manual_route_geometry_passes_through_waypoint(client):
    a, b = _two_shapes(client)
    result = client.connect_shapes(
        a, b, from_side="bottom", to_side="top", waypoints=[[5, 7]],
    )
    report = client.route_diagnostics()
    item = next(c for c in report["connectors"]
                if c["connector_id"] == result["connector_id"])
    assert item["route_kind"] == "manual"
    assert item["geometry_ok"] is True
    assert len(item["polyline"]) >= 4
    assert [5.0, 7.0] in item["polyline"]


def test_manual_route_requires_explicit_sides_and_right_angle(client):
    a, b = _two_shapes(client)
    with pytest.raises(VisioMcpError, match="explicit"):
        client.connect_shapes(a, b, waypoints=[[6, 6]])
    with pytest.raises(VisioMcpError, match="explicit"):
        client.connect_shapes(
            a, b, from_side="auto", to_side="top", waypoints=[[6, 6]],
        )
    with pytest.raises(VisioMcpError, match="right_angle"):
        client.connect_shapes(
            a, b, route="curved", from_side="bottom", to_side="top",
            waypoints=[[6, 6]],
        )


def test_waypoints_and_labels_are_validated(client):
    a, b = _two_shapes(client)
    with pytest.raises(VisioMcpError, match="at most"):
        client.connect_shapes(
            a, b, from_side="bottom", to_side="top",
            waypoints=[[i, i] for i in range(9)],
        )
    with pytest.raises(VisioMcpError, match=r"\[x, y\]"):
        client.connect_shapes(
            a, b, from_side="bottom", to_side="top", waypoints=[[1, 2, 3]],
        )
    with pytest.raises(VisioMcpError, match="label"):
        client.connect_shapes(
            a, b, from_side="bottom", to_side="top",
            waypoints=[[6, 6]], label_pos=0.5,
        )
    with pytest.raises(VisioMcpError, match="between 0 and 1"):
        client.connect_shapes(
            a, b, from_side="bottom", to_side="top", label="x", label_pos=1.5,
        )


def test_label_position_on_dynamic_connector_is_forced(client, app):
    a, b = _two_shapes(client)
    result = client.connect_shapes(
        a, b, label="L", from_side="bottom", to_side="top", label_pos=0.25,
    )
    assert result["route_kind"] == "auto"
    shape = app.ActivePage.Shapes.ItemFromID(result["connector_id"])
    assert "POINTALONGPATH" in shape.CellsU("TxtPinX").FormulaU
    assert shape.CellsU("TxtPinX").forced
    assert shape.CellsU("TxtAngle").FormulaU == "0 deg"


def test_set_connector_route_converts_dynamic_and_preserves_style(client, app):
    a, b = _two_shapes(client)
    original = client.connect_shapes(
        a, b, label="keep", from_side="bottom", to_side="top",
        line_pattern="dashed", line_weight_pt=2, line_color="#808080",
    )
    converted = client.set_connector_route(
        original["connector_id"], waypoints=[[6, 6]],
        label_pos=0.4, label_offset_in=0.2,
    )
    assert converted["replaced_id"] == original["connector_id"]
    assert converted["connector_id"] != original["connector_id"]
    assert converted["route_kind"] == "manual"
    with pytest.raises(Exception):
        app.ActivePage.Shapes.ItemFromID(original["connector_id"])
    shape = app.ActivePage.Shapes.ItemFromID(converted["connector_id"])
    assert str(shape.Text) == "keep"
    assert shape.CellsU("EndArrow").ResultIU == 5
    assert shape.CellsU("LinePattern").ResultIU == 2
    assert shape.CellsU("LineWeight").FormulaU == "2 pt"
    assert shape.CellsU("User.VisioMcpRoute").ResultIU == 1
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    assert state[converted["connector_id"]]["route_kind"] == "manual"
    assert state[converted["connector_id"]]["from_id"] == a


def test_set_connector_route_rebuilds_manual_route_in_place(client, app):
    a, b = _two_shapes(client)
    created = client.connect_shapes(
        a, b, from_side="bottom", to_side="top", waypoints=[[6, 6]],
    )
    shape = app.ActivePage.Shapes.ItemFromID(created["connector_id"])
    rows_before = int(shape.RowCount(10))
    changed = client.set_connector_route(
        created["connector_id"], waypoints=[[5, 7], [8, 5]],
    )
    assert changed["connector_id"] == created["connector_id"]
    assert changed["replaced_id"] is None
    shape = app.ActivePage.Shapes.ItemFromID(created["connector_id"])
    assert int(shape.RowCount(10)) != rows_before
    assert int(shape.RowCount(10)) >= 4


def test_set_connector_route_label_only_does_not_convert(client, app):
    a, b = _two_shapes(client)
    created = client.connect_shapes(
        a, b, label="L", from_side="bottom", to_side="top",
    )
    moved = client.set_connector_route(
        created["connector_id"], label_pos=0.75, label_offset_in=0.25,
    )
    assert moved["connector_id"] == created["connector_id"]
    assert moved["replaced_id"] is None
    assert moved["route_kind"] == "auto"
    shape = app.ActivePage.Shapes.ItemFromID(created["connector_id"])
    assert shape.CellsU("TxtPinX").forced


def test_set_connector_route_rejects_dynamic_glue(client):
    a, b = _two_shapes(client)
    created = client.connect_shapes(a, b)  # both ends dynamic
    with pytest.raises(VisioMcpError, match="pinned"):
        client.set_connector_route(created["connector_id"], waypoints=[[6, 6]])
    with pytest.raises(VisioMcpError, match="needs waypoints"):
        client.set_connector_route(created["connector_id"])


def test_manual_route_slots_are_spread_and_seen_by_state(client, app):
    a, b = _two_shapes(client)
    first = client.connect_shapes(a, b, from_side="right", to_side="left",
                                  waypoints=[])
    second = client.connect_shapes(a, b, from_side="right", to_side="left",
                                   waypoints=[])
    assert first["from_pos"] == 0.5
    assert second["from_pos"] == pytest.approx(0.3)
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    assert state[first["connector_id"]]["from_pos"] == 0.5
    assert state[second["connector_id"]]["from_pos"] == pytest.approx(0.3)


def test_manual_route_failure_rolls_back_undo_scope(client, app, monkeypatch):
    a, b = _two_shapes(client)
    def boom(shape, points):
        raise RuntimeError("injected geometry failure")
    monkeypatch.setattr(client, "_write_manual_geometry", boom)
    with pytest.raises(RuntimeError, match="injected geometry failure"):
        client.connect_shapes(
            a, b, from_side="bottom", to_side="top", waypoints=[[6, 6]],
        )
    assert app.undo_scopes[-1][1] is False


def test_set_route_options_sets_and_clears_native_lane_separation(client, app):
    client.create_document()
    on = client.set_route_options(True)
    assert (on["line_adjust_from"], on["line_adjust_to"]) == (1, 2)
    sheet = app.ActivePage.PageSheet
    assert sheet.CellsU("LineAdjustFrom").ResultIU == 1
    assert sheet.CellsU("LineAdjustTo").ResultIU == 2
    assert app.undo_scopes[-1][1] is True
    off = client.set_route_options(False)
    assert (off["line_adjust_from"], off["line_adjust_to"]) == (0, 0)
    with pytest.raises(VisioMcpError, match="page"):
        client.set_route_options(True, page_name="nope")


def test_route_diagnostics_reports_crossing_and_no_mutation(client):
    client.create_document()
    client.open_stencil("BASFLO_U.VSSX")
    ids = [s["shape_id"] for s in client.drop_shapes([
        {"master": "Process", "x": 2, "y": 8, "text": "A"},
        {"master": "Process", "x": 6, "y": 8, "text": "B"},
        {"master": "Process", "x": 4, "y": 10, "text": "C"},
        {"master": "Process", "x": 4, "y": 6, "text": "D"},
    ])["shapes"]]
    horizontal = client.connect_shapes(
        ids[0], ids[1], from_side="right", to_side="left", waypoints=[],
    )
    vertical = client.connect_shapes(
        ids[2], ids[3], from_side="bottom", to_side="top", waypoints=[],
    )
    before = len(client.get_page_state()["shapes"])
    report = client.route_diagnostics()
    assert len(client.get_page_state()["shapes"]) == before
    assert report["summary"]["crossing_count"] == 1
    by_id = {c["connector_id"]: c for c in report["connectors"]}
    assert by_id[horizontal["connector_id"]]["crossings"][0]["connector_id"] == \
        vertical["connector_id"]
    assert by_id[vertical["connector_id"]]["route_kind"] == "manual"
    assert report["summary"]["flagged_connector_ids"]


def test_route_diagnostics_reports_overlap(client):
    a, b = _two_shapes(client)
    client.connect_shapes(a, b, from_side="right", to_side="left",
                          from_pos=0.5, to_pos=0.5, waypoints=[])
    client.connect_shapes(a, b, from_side="right", to_side="left",
                          from_pos=0.5, to_pos=0.5, waypoints=[])
    report = client.route_diagnostics()
    assert report["summary"]["overlap_count"] == 1
    lengths = [entry["length_in"]
               for item in report["connectors"] for entry in item["overlaps"]]
    assert lengths and max(lengths) > 1.0
    # the reported interval is the actual shared stretch, not a whole segment
    for item in report["connectors"]:
        for entry in item["overlaps"]:
            start, end = entry["interval"]
            assert abs(abs(end[0] - start[0]) + abs(end[1] - start[1])
                       - entry["length_in"]) < 1e-6


# -- restoring Visio's automatic routing -----------------------------------

def test_restore_auto_returns_manual_route_to_a_dynamic_connector(client, app):
    a, b = _two_shapes(client)
    manual = client.connect_shapes(
        a, b, from_side="bottom", to_side="top", waypoints=[[6, 6]],
        label="keep", line_pattern="dashed", line_weight_pt=2,
        line_color="#808080",
    )
    manual_shape = app.ActivePage.Shapes.ItemFromID(manual["connector_id"])
    restored = client.set_connector_route(
        manual["connector_id"], restore_auto=True,
    )
    assert restored["restored_auto"] is True
    assert restored["route_kind"] == "auto"
    assert restored["replaced_id"] == manual["connector_id"]
    assert restored["connector_id"] != manual["connector_id"]
    with pytest.raises(Exception):
        app.ActivePage.Shapes.ItemFromID(manual["connector_id"])

    shape = app.ActivePage.Shapes.ItemFromID(restored["connector_id"])
    assert shape.OneD
    assert not client._is_manual_connector(shape)
    assert str(shape.Text) == "keep"
    assert shape.CellsU("EndArrow").ResultIU == 5
    assert shape.CellsU("LinePattern").ResultIU == 2
    assert shape.CellsU("LineWeight").FormulaU == "2 pt"
    assert shape.CellsU("LineColor").FormulaU == "RGB(128,128,128)"
    assert shape.CellsU("ShapeRouteStyle").ResultIU == 1
    assert shape.CellsU("ConLineRouteExt").ResultIU == 1
    state = {s["shape_id"]: s for s in client.get_page_state()["shapes"]}
    info = state[restored["connector_id"]]
    assert (info["from_id"], info["to_id"]) == (a, b)
    assert info["from_glue"] == info["to_glue"] == "static"
    assert (info["from_side"], info["to_side"]) == ("bottom", "top")
    assert info["route_kind"] == "auto"
    assert app.undo_scopes[-1][1] is True
    report = client.route_diagnostics()
    assert next(c for c in report["connectors"]
                if c["connector_id"] == restored["connector_id"])["route_kind"] == "auto"
    assert manual_shape not in app.ActivePage.shapes


def test_restore_auto_preserves_the_original_native_route_style(client, app):
    a, b = _two_shapes(client)
    dynamic = client.connect_shapes(
        a, b, route="curved", from_side="bottom", to_side="top",
    )
    converted = client.set_connector_route(
        dynamic["connector_id"], waypoints=[[6, 6]],
    )
    manual = app.ActivePage.Shapes.ItemFromID(converted["connector_id"])
    assert manual.CellsU("User.VisioMcpRouteStyle").ResultIU == 1
    assert manual.CellsU("User.VisioMcpRouteExt").ResultIU == 2

    restored = client.set_connector_route(
        converted["connector_id"], restore_auto=True,
    )
    shape = app.ActivePage.Shapes.ItemFromID(restored["connector_id"])
    assert shape.CellsU("ShapeRouteStyle").ResultIU == 1
    assert shape.CellsU("ConLineRouteExt").ResultIU == 2  # curved again


def test_restore_auto_on_an_automatic_connector_is_a_noop(client):
    a, b = _two_shapes(client)
    dynamic = client.connect_shapes(a, b, from_side="bottom", to_side="top")
    result = client.set_connector_route(
        dynamic["connector_id"], restore_auto=True,
    )
    assert result["connector_id"] == dynamic["connector_id"]
    assert result["replaced_id"] is None
    assert result["route_kind"] == "auto"
    assert result["restored_auto"] is False
    assert "already" in result["note"]


def test_restore_auto_rejects_waypoints(client):
    a, b = _two_shapes(client)
    manual = client.connect_shapes(
        a, b, from_side="bottom", to_side="top", waypoints=[[6, 6]],
    )
    with pytest.raises(VisioMcpError, match="cannot be combined"):
        client.set_connector_route(
            manual["connector_id"], waypoints=[[7, 5]], restore_auto=True,
        )


def test_conversion_style_failure_rolls_back_and_keeps_the_original(
        client, app, monkeypatch):
    a, b = _two_shapes(client)
    original = client.connect_shapes(
        a, b, label="keep", from_side="bottom", to_side="top",
        line_pattern="dashed",
    )
    create = client._create_manual_connector

    def create_with_guarded_style(*args, **kwargs):
        conn = create(*args, **kwargs)
        conn.CellsU("LinePattern").guarded = True  # style copy must fail
        return conn

    monkeypatch.setattr(client, "_create_manual_connector",
                        create_with_guarded_style)
    with pytest.raises(VisioMcpError, match="LinePattern"):
        client.set_connector_route(
            original["connector_id"], waypoints=[[6, 6]],
        )
    # the original connector is still there: it is deleted only after a
    # complete style copy
    shape = app.ActivePage.Shapes.ItemFromID(original["connector_id"])
    assert str(shape.Text) == "keep"
    assert shape.CellsU("LinePattern").ResultIU == 2
    assert client._is_manual_connector(shape) is False
    assert app.undo_scopes[-1][1] is False


def test_restore_auto_failure_keeps_the_manual_connector(
        client, app, monkeypatch):
    a, b = _two_shapes(client)
    manual = client.connect_shapes(
        a, b, from_side="bottom", to_side="top", waypoints=[[6, 6]],
        label="keep", line_pattern="dotted",
    )

    def boom(source, target):
        raise VisioMcpError("injected style copy failure")

    monkeypatch.setattr(client, "_copy_connector_style", boom)
    with pytest.raises(VisioMcpError, match="injected style copy failure"):
        client.set_connector_route(manual["connector_id"], restore_auto=True)
    shape = app.ActivePage.Shapes.ItemFromID(manual["connector_id"])
    assert client._is_manual_connector(shape) is True
    assert str(shape.Text) == "keep"
    assert shape.CellsU("LinePattern").ResultIU == 3
    assert app.undo_scopes[-1][1] is False


# -- diagnostics: filtering, readjustment after a node move ----------------

def test_route_diagnostics_can_filter_one_connector(client):
    a, b = _two_shapes(client)
    first = client.connect_shapes(
        a, b, from_side="right", to_side="left", waypoints=[],
        from_pos=0.5, to_pos=0.5,
    )
    second = client.connect_shapes(
        a, b, from_side="right", to_side="left", waypoints=[],
        from_pos=0.5, to_pos=0.5,
    )
    report = client.route_diagnostics(connector_id=first["connector_id"])
    assert [c["connector_id"] for c in report["connectors"]] == [
        first["connector_id"]]
    assert report["summary"]["filtered_connector_id"] == first["connector_id"]
    assert report["summary"]["connector_count"] == 1
    assert report["summary"]["overlap_count"] == 1
    assert report["connectors"][0]["overlaps"][0]["connector_id"] == \
        second["connector_id"]
    with pytest.raises(VisioMcpError, match="not a connector"):
        client.route_diagnostics(connector_id=999)


def test_manual_route_needs_readjust_after_its_node_moves(client, app):
    a, b = _two_shapes(client)
    created = client.connect_shapes(
        a, b, from_side="bottom", to_side="top", waypoints=[[6, 6]],
    )
    before = client.route_diagnostics(
        connector_id=created["connector_id"])["connectors"][0]
    assert before["orthogonal"] is True
    assert before["needs_readjust"] is False

    # Visio re-glues EndX/EndY when the target shape moves; the page-anchored
    # interior vertices stay put, so the final segment turns diagonal.
    shape = app.ActivePage.Shapes.ItemFromID(created["connector_id"])
    _move_glued_end(
        shape,
        float(shape.CellsU("EndX").ResultIU) + 1.0,
        float(shape.CellsU("EndY").ResultIU) + 0.75,
    )

    report = client.route_diagnostics(connector_id=created["connector_id"])
    item = report["connectors"][0]
    assert item["needs_readjust"] is True
    assert item["orthogonal"] is False
    assert item["non_orthogonal_segments"]
    assert item["readjust_hint"]
    assert report["summary"]["manual_needs_readjust_ids"] == [
        created["connector_id"]]
    assert report["summary"]["diagnostics_complete"] is True
