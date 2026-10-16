"""Solver-free translation and invariants; this does not execute/compile VBA.

The scalar double algorithm below mirrors the owned VBA overlay. The separate
waterline_conform module is the independent Python reference, not a solver.
"""
from __future__ import annotations

import hashlib
import importlib.util
import json
from pathlib import Path
import struct
import unittest

import numpy as np

ROOT = Path(__file__).resolve().parents[1]
REPO = ROOT.parents[1]
REFERENCE = REPO / "validation/G1R19_python26_water_20261016/python/waterline_conform.py"
spec = importlib.util.spec_from_file_location("waterline_conform_reference", REFERENCE)
reference = importlib.util.module_from_spec(spec)
spec.loader.exec_module(reference)


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def cross(xy, a, b, c):
    return ((xy[b][0]-xy[a][0])*(xy[c][1]-xy[a][1])
            -(xy[b][1]-xy[a][1])*(xy[c][0]-xy[a][0]))


def translated_split(xy, tri, ids, level, constraints=None, boundary=None):
    """Mirror WCIntersection/Clip/Triangulate and transactional plan checks."""
    oldxy = np.asarray(xy, dtype=float)
    oldtri = np.asarray(tri, dtype=np.int32)
    ids = np.asarray(ids, dtype=np.int32)
    if (oldxy.ndim != 2 or oldxy.shape[1] != 2 or oldtri.ndim != 2
            or oldtri.shape[1] != 3 or ids.shape != (len(oldtri),)
            or not np.isfinite(oldxy).all() or not np.isfinite(level)):
        raise ValueError("geometry")
    points = oldxy.tolist()
    cells, owners, shared = [], [], {}
    constraints = dict(constraints or {})
    boundary = dict(boundary or {})

    def intersection(a, b):
        if points[a][1] == level:
            return a
        if points[b][1] == level:
            return b
        a, b = sorted((a, b))
        if (a, b) not in shared:
            if not min(points[a][1], points[b][1]) < level < max(points[a][1], points[b][1]):
                raise ValueError("crossing")
            x = points[a][0] + (level-points[a][1])*(points[b][0]-points[a][0])/(points[b][1]-points[a][1])
            if not np.isfinite(x):
                raise ValueError("intersection")
            shared[a, b] = len(points)
            points.append([x, float(level)])
        return shared[a, b]

    def clip(nodes, below):
        polygon = []

        def append(n):
            if not polygon or polygon[-1] != n:
                polygon.append(n)

        for k in range(3):
            a, b = int(nodes[k]), int(nodes[(k+1) % 3])
            ina = points[a][1] <= level if below else points[a][1] >= level
            inb = points[b][1] <= level if below else points[b][1] >= level
            if ina:
                append(a)
            if ina != inb:
                append(intersection(a, b))
        if len(polygon) > 1 and polygon[-1] == polygon[0]:
            polygon.pop()
        return polygon

    def quality(a, b, c):
        denominator = ((points[a][0]-points[c][0])**2+(points[a][1]-points[c][1])**2
                       +(points[b][0]-points[a][0])**2+(points[b][1]-points[a][1])**2
                       +(points[c][0]-points[b][0])**2+(points[c][1]-points[b][1])**2)
        return cross(points, a, b, c)/denominator

    def emit(nodes, owner):
        det = cross(points, *nodes)
        if not np.isfinite(det) or det <= 1e-12:
            raise ValueError("nonpositive/degenerate")
        cells.append(tuple(nodes))
        owners.append(owner)

    def triangulate(p, owner):
        if len(p) == 3:
            emit(p, owner)
            return
        if len(p) != 4:
            raise ValueError("polygon")
        a, b, c, d = p
        if min(quality(a, b, c), quality(a, c, d)) >= min(quality(a, b, d), quality(b, c, d)):
            emit((a, b, c), owner)
            emit((a, c, d), owner)
        else:
            emit((a, b, d), owner)
            emit((b, c, d), owner)

    for e, nodes in enumerate(oldtri):
        y = oldxy[nodes, 1]
        if min(y) < level < max(y):
            triangulate(clip(nodes, True), e)
            triangulate(clip(nodes, False), e)
        else:
            emit(tuple(map(int, nodes)), e)
    xy = np.asarray(points, float)
    tri = np.asarray(cells, np.int32)
    owners = np.asarray(owners, np.int32)
    newids = ids[owners]
    # Same per-parent tolerances as the proposed VBA before committing globals.
    pa, pm = reference.triangle_moments(oldxy, oldtri)
    ca, cm = reference.triangle_moments(xy, tri)
    if np.any(pa <= 0) or np.any(ca <= 0):
        raise ValueError("orientation")
    ae = np.bincount(owners, weights=ca, minlength=len(pa))-pa
    me = np.column_stack([np.bincount(owners, weights=cm[:, j], minlength=len(pa)) for j in range(2)])-pm
    if np.any(abs(ae) > 1e-10*(1+pa)) or np.any(abs(me) > 1e-10*(1+abs(pm))):
        raise ValueError("parent moments")
    for owner, nodes in zip(owners, tri):
        a, b, c = oldtri[owner]
        det = cross(oldxy, a, b, c)
        for node in nodes:
            dx, dy = xy[node]-oldxy[a]
            v = (dx*(oldxy[c, 1]-oldxy[a, 1])-dy*(oldxy[c, 0]-oldxy[a, 0]))/det
            w = ((oldxy[b, 0]-oldxy[a, 0])*dy-(oldxy[b, 1]-oldxy[a, 1])*dx)/det
            if min(v, w, 1-v-w) < -1e-11:
                raise ValueError("containment")
    newconstraints = {}
    for nodes in tri:
        for k in range(3):
            a, b = sorted((int(nodes[k]), int(nodes[(k+1) % 3])))
            if xy[a, 1] == level and xy[b, 1] == level:
                newconstraints[a, b] = (a, b)
    for key, value in constraints.items():
        a, b = sorted(key)
        if (a, b) in shared:
            n = shared[a, b]
            newconstraints[tuple(sorted((a, n)))] = (a, n)
            newconstraints[tuple(sorted((n, b)))] = (n, b)
        else:
            newconstraints[a, b] = value
    newboundary = {}
    for key, value in boundary.items():
        a, b = sorted(key)
        if (a, b) in shared:
            n = shared[a, b]
            newboundary[tuple(sorted((a, n)))] = value
            newboundary[tuple(sorted((n, b)))] = value
        else:
            newboundary[a, b] = value
    return xy, tri, newids, owners, shared, newconstraints, newboundary


def edge_owners(tri):
    edges = {}
    for e, nodes in enumerate(tri):
        for k in range(3):
            key = tuple(sorted((int(nodes[k]), int(nodes[(k+1) % 3]))))
            edges.setdefault(key, []).append(e)
    return edges


def check_case(xy, tri, ids, level):
    boundary = {key: ("load", -1, 0., -7., 2., -3.)
                for key, owners in edge_owners(tri).items() if len(owners) == 1}
    constraints = {key: key for key, owners in edge_owners(tri).items()
                   if len(owners) == 1 or (len(owners) == 2 and ids[owners[0]] != ids[owners[1]])}
    got = translated_split(xy, tri, ids, level, constraints, boundary)
    expected = reference.split_at_waterline(xy, tri, ids, level)
    np.testing.assert_allclose(got[0], expected[0], rtol=0, atol=1e-13)
    np.testing.assert_array_equal(got[1], expected[1])
    np.testing.assert_array_equal(got[2], expected[2])
    np.testing.assert_array_equal(got[3], expected[3]["parent_element_ids"])
    stats = reference.verify_partition(xy, tri, ids, *got[:3], expected[3])
    cy = got[0][got[1], 1]
    assert not np.any((cy.min(axis=1) < level) & (cy.max(axis=1) > level))
    assert all(1 <= len(v) <= 2 for v in edge_owners(got[1]).values())
    outer = {key for key, owners in edge_owners(got[1]).items() if len(owners) == 1}
    assert outer == set(got[6])
    # Every exposed child retains the six literal physical record fields.
    assert all(record == next(iter(boundary.values())) for record in got[6].values())
    for key, owners in edge_owners(got[1]).items():
        if len(owners) == 2 and got[2][owners[0]] != got[2][owners[1]]:
            assert key in got[5]
    twice = translated_split(*got[:3], level, got[5], got[6])
    for a, b in zip(got[:3], twice[:3]):
        np.testing.assert_array_equal(a, b)
    assert len(twice[4]) == 0
    # A second geometry transform did not introduce new physical boundary data.
    assert twice[6] == got[6]
    return got, stats


def actual_d08_geometry():
    path = REPO / "versions/G1R17/results/geometry25/P02A_new_vba_mesh.bin"
    data = path.read_bytes()
    nn, ne = struct.unpack_from("<ii", data)
    xy = np.frombuffer(data, "<f8", count=nn*2, offset=8).reshape(nn, 2).copy()
    cells = np.frombuffer(data, "<i4", count=ne*4, offset=8+nn*16).reshape(ne, 4)
    assert 8+nn*16+ne*16 == len(data)
    return path, xy, cells[:, :3].copy(), cells[:, 3].copy()


def canonical_boundary(xy, tri, record):
    points = {p[0]: np.asarray(p[1:]) for p in record["points"]}
    kinds = {"固定XY": "fixed", "固定X": "roller_x", "固定Y": "roller_y", "自由": "load"}
    result = {}
    for key, owners in edge_owners(tri).items():
        if len(owners) != 1:
            continue
        midpoint = xy[list(key)].mean(axis=0)
        hits = []
        for face in record["faces"]:
            a, b = points[face[2]], points[face[3]]
            d, p = b-a, midpoint-a
            if (abs(d[0]*p[1]-d[1]*p[0]) <= 1e-10
                    and -1e-10 <= p@d <= d@d+1e-10):
                hits.append((kinds[face[5]], -1, *map(float, face[7:11])))
        if len(hits) != 1:
            raise AssertionError((key, hits))
        result[key] = hits[0]
    return result


def uniform(xy, tri, ids):
    points, children, newids, mids = xy.tolist(), [], [], {}
    for nodes, mid in zip(tri, ids):
        a, b, c = map(int, nodes)
        edgeids = []
        for u, v in ((a, b), (b, c), (c, a)):
            key = tuple(sorted((u, v)))
            if key not in mids:
                mids[key] = len(points)
                points.append(((xy[u]+xy[v])/2).tolist())
            edgeids.append(mids[key])
        ab, bc, ca = edgeids
        children.extend(((a, ab, ca), (ab, b, bc), (ca, bc, c), (ab, bc, ca)))
        newids.extend([mid]*4)
    return np.asarray(points), np.asarray(children), np.asarray(newids)


def local_centroids(xy, tri, ids):
    points, children, newids = xy.tolist(), [], []
    for e, (nodes, mid) in enumerate(zip(tri, ids)):
        if e % 3:
            children.append(tuple(nodes))
            newids.append(mid)
            continue
        a, b, c = map(int, nodes)
        center = len(points)
        points.append(xy[nodes].mean(axis=0).tolist())
        children.extend(((a, b, center), (b, c, center), (c, a, center)))
        newids.extend([mid]*3)
    return np.asarray(points), np.asarray(children), np.asarray(newids)


def body_check(xy, tri, level=3., gamma=20., gamma_water=9.81):
    points = xy[tri]
    wet = points[:, :, 1].max(axis=1) <= level
    assert not np.any((points[:, :, 1].min(axis=1) < level) & (points[:, :, 1].max(axis=1) > level))
    body = np.column_stack((np.zeros(len(tri)), -gamma+gamma_water*wet))
    maxerr = 0.
    for p, b in zip(points, body):
        pressure = gamma_water*np.maximum(level-p[:, 1], 0)
        grad = np.linalg.solve(np.column_stack((p[1]-p[0], p[2]-p[0])).T,
                               pressure[1:]-pressure[0])
        expected = np.array([0., -gamma])-grad
        maxerr = max(maxerr, float(np.max(abs(expected-b))))
    area, moment = reference.triangle_moments(xy, tri)
    return body, area @ body, np.einsum("ei,ej->ij", moment, body), maxerr


class WaterCommonTests(unittest.TestCase):
    def test_crossing_and_material_interface(self):
        xy = np.array([[0., 0.], [4., 0.], [4., 4.], [0., 4.]])
        tri = np.array([[0, 1, 2], [0, 2, 3]])
        got, _ = check_case(xy, tri, np.array([4, 9]), 2.)
        self.assertEqual(len(got[4]), 3)  # Shared diagonal counted only once.
        self.assertEqual(len(got[1]), 6)
        for key, owners in edge_owners(got[1]).items():
            if all(got[0][node, 1] == 2 for node in key) and len(owners) == 2:
                self.assertNotIn(key, got[6])  # Waterline is not an exposed face.

    def test_existing_water_vertex(self):
        xy = np.array([[0., 0.], [4., 2.], [0., 4.]])
        got, _ = check_case(xy, np.array([[0, 1, 2]]), np.array([7]), 2.)
        self.assertEqual(len(got[4]), 1)
        self.assertEqual(len(got[1]), 2)

    def test_unchanged_all_above_below_and_edge(self):
        xy = np.array([[0., 0.], [4., 0.], [0., 4.]])
        for level in (-1., 0., 4., 5.):
            got, _ = check_case(xy, np.array([[0, 1, 2]]), np.array([0]), level)
            np.testing.assert_array_equal(got[0], xy)
            self.assertEqual(len(got[4]), 0)

    def test_large_coordinate_translation(self):
        xy = np.array([[0., 0.], [4., 0.], [4., 4.], [0., 4.]])+[1e6, -1e6]
        check_case(xy, np.array([[0, 1, 2], [0, 2, 3]]), np.array([0, 0]), -999998.)

    def test_invalid_source_preserved(self):
        xy = np.array([[0., 0.], [4., 0.], [0., 4.]])
        before = xy.copy()
        for tri, level in (([[0, 2, 1]], 2.), ([[0, 0, 1]], 2.), ([[0, 1, 2]], np.nan)):
            with self.assertRaises(ValueError):
                translated_split(xy, tri, [0], level)
            np.testing.assert_array_equal(before, xy)

    def test_actual_d08_partition_and_common_load(self):
        path, xy, tri, ids = actual_d08_geometry()
        got, stats = check_case(xy, tri, ids, 3.)
        recordpath = REPO / "docs/benchmarks/G1R17_native26_20261013/data/inputs/D08A.json"
        record = json.loads(recordpath.read_text(encoding="utf-8"))
        self.assertEqual(record["source_case"]["materials"][0]["gamma"], 20.)
        old_boundary = canonical_boundary(xy, tri, record)
        literal = translated_split(xy, tri, ids, 3.,
                                   {key: key for key in old_boundary}, old_boundary)
        self.assertEqual(literal[6], canonical_boundary(*literal[:2], record))
        body, force, first, error = body_check(*got[:2])
        np.testing.assert_allclose(force, [0., -2571.4], atol=1e-9, rtol=0)
        self.assertLess(error, 1e-10)
        refined = uniform(*got[:3])
        _, newforce, newfirst, newerror = body_check(*refined[:2])
        np.testing.assert_allclose(force, newforce, atol=1e-9, rtol=0)
        np.testing.assert_allclose(first, newfirst, atol=1e-8, rtol=0)
        self.assertLess(newerror, 1e-9)
        local = local_centroids(*got[:3])
        _, localforce, localfirst, localerror = body_check(*local[:2])
        np.testing.assert_allclose(force, localforce, atol=1e-9, rtol=0)
        np.testing.assert_allclose(first, localfirst, atol=1e-8, rtol=0)
        self.assertLess(localerror, 1e-9)
        output = ROOT / "results"
        output.mkdir(exist_ok=True)
        fixture = output / "D08A_expected_common_geometry.npz"
        np.savez_compressed(fixture, xy=got[0], tri=got[1], material_id=got[2],
                            parent_element_id=got[3], body_total=body,
                            boundary_edges=np.asarray(list(literal[6]), dtype=np.int32),
                            boundary_records=np.asarray(list(literal[6].values()), dtype=str),
                            waterline_constraints=np.asarray([key for key in got[5] if all(got[0][n, 1] == 3 for n in key)], dtype=np.int32))
        report = dict(status="PASS", scope="solver-free scalar VBA algorithm translation versus independent Python reference; VBA execution/compile not tested",
                      parent_geometry=str(path.relative_to(REPO)), parent_sha256=sha(path),
                      reference_sha256=sha(REFERENCE), elements_before=len(tri),
                      elements_after=len(got[1]), nodes_before=len(xy), nodes_after=len(got[0]),
                      crossing_shared_edges=len(got[4]), uniform_elements=len(refined[1]), local_centroid_elements=len(local[1]),
                      literal_canonical_boundary_records_preserved=True,
                      canonical_input_sha256=sha(recordpath),
                      total_force=force.tolist(), first_moment=first.tolist(),
                      max_nodal_P1_body_error=error, uniform_max_nodal_P1_body_error=newerror,
                      local_max_nodal_P1_body_error=localerror,
                      geometry_invariants=stats, fixture_sha256=sha(fixture))
        (output / "STATIC_GEOMETRY_COMPARISON.json").write_text(json.dumps(report, ensure_ascii=False, indent=2)+"\n", encoding="utf-8")

    def test_source_guards_and_no_added_physical_face(self):
        mesh = (ROOT / "src/vba/RPX_Mesh.bas").read_text(encoding="utf-8")
        water = (ROOT / "src/vba/RPX_Water.bas").read_text(encoding="utf-8")
        section = mesh[mesh.index("' Horizontal prescribed water"):mesh.index("Private Sub CavityEdge")]
        self.assertNotIn("DFaces", section)
        for text in ("water_common_plan", "water_common_commit", "WCBoundaryEqual(expectedBoundary)",
                     "WATER_COMMON_PARENT_CONTAINMENT", "WATER_COMMON_PARENT_MOMENTS",
                     "RPX_ResetPreparedState", "RPX_LowerCacheReset", "RPX_FsBracketReset",
                     "WCInterrupted(errorNumber, errorText)", "WATER_COMMON_ROLLBACK_FAILED"):
            self.assertIn(text, section)
        for text in ("RPX_SeepMode(ws) <> 2", "SuctionOn()", "RR <> 0", "If y <> level", "CDbl(values(i, 1)) <> level",
                     "RPX_WaterCommonConforming(level)", "If below Then wy = commonGamma"):
            self.assertIn(text, water)
        self.assertLess(max(map(len, mesh.splitlines())), 1023)
        self.assertLess(max(map(len, water.splitlines())), 1023)


if __name__ == "__main__":
    unittest.main(verbosity=2)
