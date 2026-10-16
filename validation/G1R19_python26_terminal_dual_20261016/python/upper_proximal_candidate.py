"""Recover a q=8 upper primal with a capped minimum-change program.

The quadratic objective is an auxiliary feasibility objective. Its dual is
never evaluated as an original optimization dual or used to finish a root.
Only a fresh audit of the unchanged original program can publish a field.
"""
from pathlib import Path
import argparse
import hashlib
import json
import shutil
import time

import accuracy_driver as driver
import experiment_core as core

np, sparse, clarabel = core.np, core.sparse, core.clarabel


def file_hash(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def program_hash(program):
    digest = hashlib.sha256()
    for value in program:
        if sparse.issparse(value):
            arrays = (np.asarray(value.shape, dtype='<i8'), value.indptr,
                      value.indices, value.data)
        else:
            arrays = (np.asarray(value),)
        for array in arrays:
            digest.update(str(array.dtype).encode())
            digest.update(np.asarray(array.shape, dtype='<i8').tobytes())
            digest.update(np.ascontiguousarray(array).tobytes())
    return digest.hexdigest()


def solve_candidate(prepared, center, alpha=.01, seconds=600., cap=None, cone_scaling=False):
    """Return an original primal audit; solver status supplies no proof."""
    cap = core.UPPER_TARGET - 1e-6 if cap is None else float(cap)
    if (prepared.mode != 'upper' or prepared.q != 8 or prepared.kind != 'none'
            or not np.isfinite([alpha, seconds, cap]).all()
            or alpha <= 0 or seconds <= 0
            or cap > core.UPPER_TARGET - 1e-6):
        raise ValueError('An original q=8 upper program, positive budgets and strict cap are required')
    E, b, G, h, dims, cost = prepared.original
    if (not np.array_equal(cost, prepared.physical_original_objective)
            or (E.shape, G.shape, len(b), len(h), tuple(dims))
            != prepared.physical_original_layout):
        raise ValueError('The original physical objective and layout are required')
    center = np.asarray(center, dtype=float)
    if center.shape != cost.shape or not np.isfinite(center).all():
        raise ValueError('The center must be a finite original-layout vector')
    before = program_hash(prepared.original)
    matrix = sparse.vstack([E, -G, sparse.csc_matrix(cost.reshape(1, -1))], format='csc')
    rhs = np.r_[b, h, cap]
    cones = prepared.cones + [clarabel.NonnegativeConeT(1)]
    row_scales = np.r_[np.ones(len(b)), core.cone_row_scales(G, dims), 1.] if cone_scaling else np.ones(len(rhs))
    # Each cone receives a single positive multiplier, including all tails.
    # The extra objective cap and original equality rows keep multiplier 1.
    if cone_scaling:
        matrix = (sparse.diags(row_scales) @ matrix).tocsc()
        rhs = rhs * row_scales
    config = dict(driver.ACCURATE, time_limit=float(seconds), tol_gap_abs=1e-13,
                  tol_gap_rel=1e-13, tol_feas=1e-13,
                  iterative_refinement_max_iter=30,
                  iterative_refinement_abstol=1e-15,
                  iterative_refinement_reltol=1e-15)
    solver = clarabel.DefaultSolver(sparse.eye(len(cost), format='csc') * alpha,
                                   -alpha * center, matrix, rhs, cones,
                                   core.settings(config, 350))
    started = time.perf_counter()
    solution = solver.solve()
    elapsed = time.perf_counter() - started
    field = np.asarray(solution.x).copy()
    # In particular, do not supply solution.z to core.evaluate.
    audit = core.evaluate(prepared, field)
    after = program_hash(prepared.original)
    if before != after:
        raise RuntimeError('The original physical program changed')
    return field, np.asarray(solution.z).copy(), {
        'status': str(solution.status), 'iterations': solution.iterations,
        'seconds': elapsed, 'alpha': float(alpha), 'cap': cap, 'config': config,
        'original_program_sha256_before': before,
        'original_program_sha256_after': after, 'original_audit': audit,
        'common_positive_cone_scaling': bool(cone_scaling),
        'row_scales_sha256': hashlib.sha256(row_scales.tobytes()).hexdigest(),
        'root_complete': False, 'modified_dual_is_original_root_guide': False}


def load_center(case, seed):
    _, record = driver.load_model(case)
    stored = (seed / 'best_upper_certificate.json').is_file()
    names = (['best_upper_certificate.json', 'best_upper_mesh.npz', 'best_upper_field.npz']
             if stored else ['audit.json', 'original_mesh.npz', 'original_candidate.npz'])
    paths = [seed / name for name in names]
    before = {name: file_hash(path) for name, path in zip(names, paths)}
    metadata = json.loads(paths[0].read_text(encoding='utf8'))
    mesh = driver.load_saved_mesh(paths[1], record)
    with np.load(paths[2], allow_pickle=False) as data:
        if (data['mode'].shape != () or str(data['mode']) != 'upper'
                or data['q'].shape != () or data['q'].item() != 8
                or data['fs'].shape != ()):
            raise ValueError('The source must use the original q=8 upper layout')
        source_fs = float(data['fs'])
        center = data['x'].copy()
        for key, expected in [('case', case), ('mesh_id', mesh.mesh_id),
                              ('physical_model_id', mesh.physical_model_id)]:
            if key in data.files and str(data[key]) != expected:
                raise ValueError(f'Source vector {key} mismatch')
        reference_fs = metadata.get('reference_witness_fs', source_fs)
        if ('reference_witness_fs' in data.files
                and float(data['reference_witness_fs']) != reference_fs):
            raise ValueError('Reference strength metadata mismatch')
    if (metadata.get('case') != case or metadata.get('q') != 8
            or metadata.get('mode', 'upper') != 'upper'
            or metadata.get('fs') != source_fs
            or metadata.get('mesh_id') != mesh.mesh_id
            or metadata.get('physical_model_id') != mesh.physical_model_id
            or not np.isfinite([source_fs, reference_fs]).all()
            or min(source_fs, reference_fs) <= 0):
        raise ValueError('Source case, strength, mesh or physical model mismatch')
    reference = core.Prepared(mesh, reference_fs, 'upper', 8, reduction='none')
    if stored and reference_fs != source_fs:
        raise ValueError('A stored certificate must be audited at its own strength')
    if center.shape != reference.original[-1].shape or not np.isfinite(center).all():
        raise ValueError('Invalid original source vector')
    reference_audit = core.evaluate(reference, center)
    if stored and not (reference_audit['physical_pass'] and reference_audit['witness']):
        raise ValueError('A stored certificate fails the fresh original physical gate')
    if before != {name: file_hash(path) for name, path in zip(names, paths)}:
        raise ValueError('Source files changed during loading')
    return mesh, center, source_fs, {
        'kind': 'stored_upper' if stored else 'raw_original_upper_center_only',
        'reference_fs': float(reference_fs), 'reference_original_audit': reference_audit,
        'source_fs': source_fs, 'source_file_sha256': before,
        'source_flags_ignored': True}, paths


def run(args):
    if args.output.exists():
        raise ValueError('The output must be a new directory')
    candidate_source = Path(__file__).read_bytes()
    candidate_hash = hashlib.sha256(candidate_source).hexdigest()
    implementation_sources = {name: (driver.REPO / name).read_bytes()
                              for name in driver.SOURCE_HASHES}
    implementation_hashes = {name: hashlib.sha256(value).hexdigest()
                             for name, value in implementation_sources.items()}
    if implementation_hashes != driver.SOURCE_HASHES:
        raise ValueError('Implementation changed after importing the experiment')
    mesh, center, source_fs, provenance, paths = load_center(args.case, args.seed)
    fs = source_fs if args.fs is None else float(args.fs)
    if not np.isfinite(fs) or fs <= 0:
        raise ValueError('A finite positive target strength is required')
    prepared = core.Prepared(mesh, fs, 'upper', 8, reduction='none')
    target_center_audit = core.evaluate(prepared, center)
    field, modified_dual, report = solve_candidate(prepared, center, args.alpha,
                                                  args.seconds, args.cap,
                                                  getattr(args, 'cone_scaling', False))
    if provenance['source_file_sha256'] != {p.name: file_hash(p) for p in paths}:
        raise ValueError('Source files changed before publication')
    if (candidate_hash != file_hash(Path(__file__))
            or implementation_hashes != {name: file_hash(driver.REPO / name)
                                         for name in implementation_sources}):
        raise ValueError('Experiment implementation changed during the solve')
    args.output.mkdir(parents=True)
    sources = args.output / 'source_files'
    sources.mkdir()
    for path in paths:
        shutil.copyfile(path, sources / path.name)
    if (provenance['source_file_sha256'] != {p.name: file_hash(p) for p in paths}
            or provenance['source_file_sha256']
            != {p.name: file_hash(sources / p.name) for p in paths}):
        raise ValueError('Copied source files differ from the frozen inputs')
    (args.output / 'experiment_source.py').write_bytes(candidate_source)
    for name, content in implementation_sources.items():
        path = args.output / 'implementation_sources' / name
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(content)
    driver.save_mesh(args.output / 'original_mesh.npz', mesh)
    np.savez_compressed(args.output / 'source_candidate.npz', x=center,
                        fs=source_fs, mode='upper', q=8)
    # Deliberately use no 'z' key: this is not an original optimization dual.
    np.savez_compressed(args.output / 'original_candidate.npz', x=field,
                        modified_dual=modified_dual, fs=fs, mode='upper', q=8)
    audit = report['original_audit']
    audit.update(case=args.case, mode='upper', q=8, fs=fs,
                 mesh_id=mesh.mesh_id, physical_model_id=mesh.physical_model_id)
    report.update(case=args.case, fs=fs, q=8, mesh_id=mesh.mesh_id,
                  physical_model_id=mesh.physical_model_id,
                  source=str(args.seed.resolve()), source_provenance=provenance,
                  target_center_original_audit=target_center_audit,
                  implementation_sha256=implementation_hashes,
                  candidate_code_sha256=candidate_hash)
    if audit['physical_pass'] and audit['witness']:
        certificate = driver.Certificate('upper', fs, mesh, field, audit,
                                         'Capped proximal QP, fresh original q=8 audit', False)
        core.write(args.output / 'import_audit.json', {
            'case': args.case, 'certificates': {'upper': certificate.publish(args.output)}})
    core.write(args.output / 'audit.json', report)
    print('UPPER_CAP_QP', args.case, report['status'], audit['value'],
          audit['physical_pass'], audit['witness'], report['seconds'], flush=True)
    return report


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--case', required=True)
    parser.add_argument('--seed', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--fs', type=float)
    parser.add_argument('--alpha', type=float, default=.01)
    parser.add_argument('--seconds', type=float, default=600.)
    parser.add_argument('--cap', type=float, default=core.UPPER_TARGET - 1e-6)
    parser.add_argument('--cone-scaling', action='store_true')
    run(parser.parse_args())
