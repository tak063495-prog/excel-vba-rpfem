"""Optional angular enrichment for independently retained lower witnesses."""
from __future__ import annotations
import experiment_core as core
from experiment_core import np, topology


def lower_stress_score(mesh, field, fs, with_metadata=False, proximity_background=.1):
    """Rank lower-mesh refinements by stress variation near the yield surface.

    This is a mesh-selection heuristic, not an error estimate or certificate.
    It combines variation of the six quadratic Bernstein controls with yield
    proximity and same-material stress jumps. New bounds still require the
    full original physical audit and their own root search.
    """
    field = np.asarray(field, dtype=float)
    if field.shape != (18 * len(mesh.tri) + 1,) or not np.isfinite(field).all():
        raise ValueError('A finite lower stress field on this mesh is required')
    if not np.isfinite(fs) or fs <= 0:
        raise ValueError('A finite positive strength factor is required')
    if not np.isfinite(proximity_background) or not 0 <= proximity_background <= 1:
        raise ValueError('The yield-proximity background must be finite and between zero and one')
    sigma = field[:-1].reshape(-1, 6, 3)
    area, _, edges = topology(mesh)
    reference = max(float(np.sqrt(np.mean(np.sum(sigma*sigma, axis=2)))), 1e-12)
    variation = np.mean(np.sum((sigma - sigma.mean(axis=1)[:, None])**2, axis=2), axis=1) / reference**2
    proximity = np.zeros(len(mesh.tri))
    for e, values in enumerate(sigma):
        cohesion, phi = core.audit_strength(mesh, e, fs)
        cap = 2 * cohesion * np.cos(phi)
        margin = cap - np.hypot(values[:, 0] - values[:, 1], 2*values[:, 2]) - np.sin(phi)*(values[:, 0]+values[:, 1])
        proximity[e] = np.mean(np.exp(-8*np.maximum(margin, 0)/(reference+cap)))
    proximity_weight = proximity_background + (1-proximity_background)*proximity
    score = area * variation * proximity_weight
    edge_score = np.zeros(len(mesh.tri))
    for edge in edges:
        e, f = edge.elements
        if f < 0 or mesh.material_id[e] != mesh.material_id[f]:
            continue
        jump = sigma[e, edge.loc[0]] - sigma[f, edge.loc[1]]
        variation_jump = float(np.mean(np.sum(jump*jump, axis=1))) / reference**2
        term = .25 * (area[e]+area[f])/2 * variation_jump * (proximity_weight[e]+proximity_weight[f])/2
        edge_score[e] += term/2
        edge_score[f] += term/2
    score += edge_score
    if not np.isfinite(score).all():
        raise ValueError('Nonfinite stress refinement score')
    if with_metadata:
        return score, {'method': 'Bernstein_stress_variation_yield_proximity_and_same_material_jumps',
                       'certificate': False, 'fs': float(fs), 'stress_reference': reference,
                       'proximity_background': float(proximity_background),
                       'variation': variation.tolist(), 'yield_proximity': proximity.tolist(),
                       'edge_score': edge_score.tolist(), 'score': score.tolist()}
    return score


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
