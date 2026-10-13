"""Optional angular enrichment for independently retained lower witnesses."""
from __future__ import annotations
import experiment_core as core
from experiment_core import np, topology


def angular_fan_refine(mesh, rec, scope='star'):
    """Split triangles incident to physical zero-stress c=0 boundary vertices.

    Centroids add angular sectors without moving an existing edge, material
    interface, or exterior vertex. Each mapping entry indexes the input mesh,
    allowing polynomial witness prolongation directly from that input mesh.
    The returned mesh is a candidate; it becomes a bound only after auditing.
    """
    if scope not in ('star', 'c0'):
        raise ValueError('scope must be star or c0')
    _, _, edges = topology(mesh)
    c0 = mesh.materials[mesh.material_id, 0] == 0
    vertices = set()
    for edge in edges:
        e, f = edge.elements
        if f < 0 and c0[e] and edge.kind == 'load' and np.all(edge.loads == 0):
            vertices.update((int(edge.a), int(edge.b)))
    selected = {e for e, tri in enumerate(mesh.tri)
                if (scope == 'star' or c0[e]) and any(int(n) in vertices for n in tri)}
    xy = mesh.xy.tolist()
    tri, ids, parent = [], [], []
    for e, (nodes, material) in enumerate(zip(mesh.tri, mesh.material_id)):
        a, b, c = map(int, nodes)
        if e in selected:
            center = len(xy)
            xy.append(mesh.xy[nodes].mean(axis=0).tolist())
            children = [(a, b, center), (b, c, center), (c, a, center)]
        else:
            children = [(a, b, c)]
        tri.extend(children)
        ids.extend([int(material)] * len(children))
        parent.extend([e] * len(children))
    out = core.mesh_from_arrays(np.asarray(xy), np.asarray(tri), np.asarray(ids), rec)
    out.provenance['parent_element_ids'] = parent
    return out, {'method': 'c0_zero_vertex_angular_fan', 'scope': scope,
                 'zero_vertices': sorted(vertices), 'split_elements': len(selected),
                 'elements_before': len(mesh.tri), 'elements_after': len(out.tri)}
