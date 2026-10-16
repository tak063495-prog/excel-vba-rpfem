"""Build and audit a thin upper wedge for a homogeneous, zero-cohesion slope.

This optional candidate preserves the original six-vertex polygon, material,
body force and exterior conditions. It changes the upper mesh only. An
analytical wedge threshold initializes the search; only an original q8 primal
audit publishes a bound, and both original signed endpoints must close its root.
"""
from __future__ import annotations
import argparse
import hashlib
import math
from pathlib import Path
from scipy.optimize import brentq
import accuracy_driver as driver
import experiment_core as core
from model_support.upper_assembly import strength


def wedge_mesh(record, depth):
    """Require the supported monotone polygon and a single c=0 soil region."""
    if not driver.finite(depth) or depth <= 0:
        raise ValueError('A finite positive wedge depth is required')
    source = record['source_case']
    materials = core.np.asarray(record['materials'], dtype=float)
    outline = core.np.asarray(source['outline'], dtype=float)
    if (source.get('pore') or len(materials) != 1 or materials.shape != (1, 5)
            or materials[0, 1] != 0 or not 0 < materials[0, 2] < 90
            or len(source.get('regions', [])) != 1
            or not core.np.array_equal(core.np.asarray(source['regions'][0]['xy']), outline)):
        raise ValueError('A dry, single-region zero-cohesion frictional soil is required')
    if record.get('settings', {}).get('SEISMIC_KH', 0) != 0:
        raise ValueError('The original dry reconstruction does not support seismic loading')
    if any(any(v != 0 for v in face[key]) for face in record['boundary_explicit']
           for key in ['load0', 'load1']):
        raise ValueError('This wedge candidate requires gravity-only exterior loading')
    if outline.shape != (6, 2) or not core.np.isfinite(outline).all():
        raise ValueError('A finite six-vertex monotone slope polygon is required')
    p0, p1, p2, p3, p4, p5 = outline
    xleft, xright, xtoe, xcrest = p0[0], p1[0], p3[0], p4[0]
    floor, toe, crest = p0[1], p2[1], p4[1]
    supported = (xleft < xcrest < xtoe < xright and floor < toe < crest
        and p1[1] == floor and p2[0] == xright and p3[1] == toe
        and p5[0] == xleft and p5[1] == crest)
    if not supported or not depth < crest - toe:
        raise ValueError('Unsupported polygon orientation or wedge depth')
    xy = core.np.vstack([outline, [xcrest, crest-depth], [xcrest, (floor+toe)/2]])
    tri = core.np.array([[6,3,4],[0,1,7],[1,2,7],[2,3,7],[7,3,6],
                         [7,6,5],[6,4,5],[0,7,5]], dtype=core.np.int32)
    mesh = core.mesh_from_arrays(xy, tri, core.np.zeros(8, core.np.int32), record)
    area, _, edges = core.topology(mesh)
    # mesh_from_arrays checks each original exterior condition and region area.
    if abs(float(area.sum())-source['area']) > 1e-9:
        raise ValueError('The candidate changed the original area')
    if len([edge for edge in edges if edge.elements[1] < 0]) != 6:
        raise ValueError('The candidate changed the original polygon edges')
    angle = math.atan((crest-toe-depth)/(xtoe-xcrest))
    change = {'method': 'homogeneous_c0_thin_slope_wedge', 'depth': depth,
        'elements': 8, 'same_physical_model': True, 'source_outline': outline.tolist(),
        'source_materials': record['materials'], 'original_area': source['area'],
        'candidate_area': float(area.sum()), 'boundary_conditions_reconstructed_from_original': True,
        'physical_model_id': mesh.physical_model_id, 'mesh_id': mesh.mesh_id}
    return mesh, angle, change


def explicit_witness(mesh, angle, fs):
    """Audit the constant moving triangle; no analytical bound is inferred."""
    phi = strength(mesh, 0, fs)
    velocity = core.np.zeros((len(mesh.tri), 6, 2))
    velocity[0,:,0] = 1.
    velocity[0,:,1] = -math.tan(angle-phi)
    candidate = velocity.ravel()
    prepared = core.Prepared(mesh, fs, 'upper', 8, reduction='none')
    E, b, *_ = prepared.original
    work_rows = core.np.flatnonzero(b != 0)
    if len(work_rows) != 1:
        raise ValueError('Unexpected original reference-work layout')
    row = int(work_rows[0]);work = float((E[row]@candidate)[0]/b[row])
    if not driver.finite(work) or work <= 0:
        raise ValueError('The explicit wedge has no positive reference work')
    candidate /= work
    audit = core.evaluate(prepared, candidate)
    audit.update(case=mesh.provenance['case'], mode='upper', q=8, fs=fs,
        mesh_id=mesh.mesh_id, physical_model_id=mesh.physical_model_id,
        origin='explicit_constant_velocity_wedge_full_original_primal_audit', work_normalization=work)
    return candidate, audit


def run(args):
    if (not driver.finite(args.depth, args.tolerance, args.solve_seconds)
            or args.depth <= 0 or not 0 < args.tolerance <= 1e-4 or args.solve_seconds <= 0):
        raise ValueError('Depth, final Fs tolerance and solve budget must be valid')
    _, record = driver.load_model(args.case)
    mesh, angle, change = wedge_mesh(record, args.depth)
    critical = brentq(lambda fs: strength(mesh, 0, fs)-angle, 1e-4, 1000., xtol=1e-13)
    change['analytical_wedge_threshold_candidate_only'] = critical
    change['analytical_threshold_is_a_published_bound'] = False
    folder = args.output.resolve()/args.case;folder.mkdir(parents=True, exist_ok=True)
    driver.save_mesh(folder/'candidate_mesh.npz', mesh)
    core.write(folder/'mesh_change.json', change)
    candidate, audit = explicit_witness(mesh, angle, critical+5e-5)
    core.np.savez_compressed(folder/'explicit_wedge_candidate.npz', x=candidate,
        fs=audit['fs'], mode='upper', q=8)
    core.write(folder/'explicit_wedge_audit.json', audit)
    warm = None
    if audit['witness']:
        audit['search_sign'] = -1
        warm = driver.Certificate('upper', audit['fs'], mesh, candidate, audit,
            'independently audited explicit homogeneous c0 wedge', False)
        warm.publish(folder)
    certified, root = driver.search(mesh, 'upper', (max(1e-5, critical-.002), audit['fs']),
        folder, args.case+'_wedge_upper', args.tolerance, args.solve_seconds, warm)
    source_hashes = dict(driver.SOURCE_HASHES)
    source_hashes[str(Path(__file__).resolve().relative_to(driver.REPO))] = hashlib.sha256(Path(__file__).read_bytes()).hexdigest()
    report = {'case': args.case, 'upper_only': True, 'one_percent_target_not_evaluated': True,
        'root': root, 'mesh_change': change, 'implementation_hashes': source_hashes}
    if certified is not None:
        metadata = certified.publish(folder);report['upper'] = metadata
        core.write(folder/'import_audit.json', {'case': args.case, 'certificates': {'upper': metadata}})
    core.write(folder/'summary.json', report)
    print('WEDGE_FINAL', args.case, certified.fs if certified else None,
          certified.root_complete if certified else False, flush=True)
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--case', required=True)
    parser.add_argument('--depth', type=float, default=.01)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--tolerance', type=float, default=1e-4)
    parser.add_argument('--solve-seconds', type=float, default=30.)
    run(parser.parse_args())
