"""Optional generic mesh choices and stable candidates; never certificates."""
from __future__ import annotations
import numpy as np
import experiment_core as core
from precision_interpolation import local_barycentric, precision_prolong_upper


def stable_lower_prolong(parent, child, field):
    field = np.asarray(field, dtype=float)
    ancestry = np.asarray(child.provenance.get('parent_element_ids'), dtype=int)
    if (field.shape != (18*len(parent.tri)+1,) or not np.isfinite(field).all()
            or ancestry.shape != (len(child.tri),) or np.any(ancestry < 0)
            or np.any(ancestry >= len(parent.tri))
            or parent.physical_model_id != child.physical_model_id):
        raise ValueError('Finite same-model nested lower field and ancestry required')
    old = field[:-1].reshape(-1, 6, 3)
    result = []
    for nodes, e in zip(child.tri, ancestry):
        vertices = child.xy[nodes]
        oldvertices = parent.xy[parent.tri[e]]
        if np.array_equal(vertices, oldvertices):
            result.append(old[e].copy())
            continue
        points = np.concatenate([vertices, (vertices+np.roll(vertices, -1, axis=0))/2])
        bary = local_barycentric(oldvertices, points)
        if not np.isfinite(bary).all() or np.min(bary) < -2e-10 or np.max(bary) > 1+2e-10:
            raise ValueError('Lower child leaves its parent triangle')
        a, b, c = bary.T
        basis = np.column_stack([b*b, c*c, 2*a*b, 2*b*c, 2*c*a])
        values = np.asarray(old[e][0]+basis@(np.asarray(old[e][1:], dtype=np.longdouble)-old[e][0]), dtype=float)
        values[3:] = 2*values[3:]-.5*(values[:3]+np.roll(values[:3], -1, axis=0))
        result.append(values)
    return np.r_[np.asarray(result).ravel(), field[-1]]


def stable_prolong(parent, child, field, mode, q=8):
    if mode == 'upper':
        return precision_prolong_upper(parent, child, field, q)
    if mode == 'lower':
        return stable_lower_prolong(parent, child, field)
    raise ValueError('Expected lower or upper')


def upper_refinement_score(mesh, field, fs, kinematic_fraction=.2):
    """Mix actual dissipative work and the existing kinematic indicator.

    The c=0 layer can dominate the velocity norm while contributing no
    objective. Preserve a positive kinematic share but allocate more degrees
    of freedom to the cohesive dissipation. Check the work decomposition
    against the full unmodified original objective. Phi=0 uses the existing
    kinematic rule rather than substituting a divergent c/tan(phi) formula.
    """
    field = np.asarray(field, dtype=float)
    if (field.ndim != 1 or len(field) < 12*len(mesh.tri) or not np.isfinite(field).all()
            or not np.isfinite(fs) or fs <= 0 or not np.isfinite(kinematic_fraction)
            or not 0 <= kinematic_fraction <= 1):
        raise ValueError('Finite field, strength and convex indicator weight required')
    velocity = core.indicator(mesh, field)
    if (not np.isfinite(velocity).all() or np.any(velocity < 0)):
        raise ValueError('Invalid original kinematic score')
    if np.any(mesh.materials[mesh.material_id, 1] == 0):
        return velocity, dict(method='existing_kinematic_phi0_fallback', certificate=False)
    from model_support.upper_assembly import strength, reduced_cohesion
    from model_support.p2 import gradients
    values = field[:12*len(mesh.tri)].reshape(-1, 6, 2)
    area, grad, edges = core.topology(mesh)
    strain, jump = np.zeros(len(mesh.tri)), np.zeros(len(mesh.tri))
    fixed = 0.
    for e in range(len(mesh.tri)):
        factor = reduced_cohesion(mesh, e, fs)/np.tan(strength(mesh, e, fs))
        for k in range(3):
            derivative = values[e].T@gradients(np.eye(3)[k], grad[e])
            strain[e] += area[e]*factor*np.trace(derivative)/3
        fixed += float(area[e]/3*np.sum(values[e, 3:]*mesh.body[e, :, 0]))/mesh.stress_scale
    for edge in edges:
        e, f = edge.elements
        if f < 0:
            fixed += float(edge.length/6*np.array([1., 1., 4.])@(values[e, edge.loc[0]]@edge.loads[:, 0]))/mesh.stress_scale
        elif mesh.material_id[e] == mesh.material_id[f]:
            factor = reduced_cohesion(mesh, e, fs)/np.tan(strength(mesh, e, fs))
            opening = (values[f, edge.loc[1]]-values[e, edge.loc[0]])@edge.normal
            contribution = edge.length*factor*(np.array([1., 1., 4.])@opening)/6
            jump[e] += contribution/2
            jump[f] += contribution/2
    signed = strain+jump
    original = core.program(mesh, fs, 'upper', 8)
    objective = float(original[-1]@field)
    error = abs(float(signed.sum()-fixed)-objective)
    if not np.isfinite(signed).all() or not np.isfinite(objective) or error > 1e-10*(1+abs(objective)):
        raise ValueError('Dissipation decomposition differs from original objective')
    dissipative = np.maximum(signed, 0)
    if dissipative.sum() <= 0 or velocity.sum() <= 0:
        return velocity, dict(method='existing_kinematic_zero_dissipation_fallback', certificate=False)
    score = ((1-kinematic_fraction)*dissipative/dissipative.sum()
             +kinematic_fraction*velocity/velocity.sum())
    return score, dict(method='original_dissipation_plus_kinematic', certificate=False,
        kinematic_fraction=kinematic_fraction, objective=objective,
        decomposition_error=error, minimum_signed_cell=float(signed.min()),
        per_material={int(mid): dict(elements=int(np.sum(mesh.material_id == mid)),
            dissipation=float(signed[mesh.material_id == mid].sum()),
            kinematic=float(velocity[mesh.material_id == mid].sum()),
            mixed_score=float(score[mesh.material_id == mid].sum())) for mid in np.unique(mesh.material_id)})
