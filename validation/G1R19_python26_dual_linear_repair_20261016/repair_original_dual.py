"""One bounded linear correction of free original dual y; no conic solve."""
from pathlib import Path
import argparse
import hashlib
import json
import os
import sys
import time

for key in ('OPENBLAS_NUM_THREADS', 'OMP_NUM_THREADS', 'MKL_NUM_THREADS'):
    os.environ[key] = '1'
ROOT = Path(__file__).resolve().parent
PIPELINE = ROOT.parent / 'G1R19_python26_terminal_final_20261016'
sys.path.insert(0, str(PIPELINE / 'python'))
import accuracy_driver as driver
import report_accuracy as report
from original_dual import audit_original_dual
from scipy.sparse.linalg import lsmr


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def run(args):
    started = time.perf_counter()
    if args.output.exists():
        raise ValueError('A fresh output directory is required')
    args.output.mkdir(parents=True)
    sources = driver.source_hashes()
    sources[str(Path(__file__).resolve())] = sha(__file__)
    _, rec = driver.load_model(args.case)
    candidate = report.collect_candidate(args.source / 'best_upper_certificate.json', 'upper',
        args.case, rec, [args.source.resolve()], 1e-4)
    if candidate is None:
        raise ValueError('An unchanged full original physical upper witness is required')
    tracked = dict(candidate['tracked'])
    tracked[args.dual] = sha(args.dual)
    with driver.np.load(args.dual, allow_pickle=False) as data:
        dual = data['dual'].copy()
        fs = float(data['fs'])
        if int(data['q']) != 8 or str(data['mode']) != 'upper':
            raise ValueError('Original q8 upper dual required')
    if not 0 < candidate['fs'] - fs <= 1e-4 * max(1., candidate['fs']):
        raise ValueError('The guide must be below this same-mesh physical endpoint within tau')
    mesh = candidate['mesh']
    prepared = driver.core.Prepared(mesh, fs, 'upper', 8, reduction='none')
    E, b, G, h, dims, cost = prepared.original
    if dual.shape != (len(b) + len(h),) or not driver.np.isfinite(dual).all():
        raise ValueError('Invalid original dual layout')
    before = audit_original_dual(prepared, dual)
    y, z = dual[:len(b)].copy(), dual[len(b):].copy()
    # The convention is c + E.T*y - G.T*z = 0. Only y is unconstrained.
    rhs = -(cost + E.T @ y - G.T @ z)
    linear_started = time.perf_counter()
    correction = lsmr(E.T.tocsr(), rhs, atol=1e-15, btol=1e-15,
                      conlim=1e12, maxiter=200)
    linear_seconds = time.perf_counter() - linear_started
    repaired = driver.np.r_[y + correction[0], z]
    if not driver.np.array_equal(repaired[len(b):], z):
        raise AssertionError('The cone variable must remain bitwise unchanged')
    after = audit_original_dual(prepared, repaired)
    within_budget = linear_seconds <= 10.
    accepted = bool(after['accepted'] and within_budget)
    driver.np.savez_compressed(args.output / 'repaired_original_dual.npz', dual=repaired,
        fs=fs, mode='upper', q=8, case=args.case, mesh_id=mesh.mesh_id,
        physical_model_id=mesh.physical_model_id, policy=mesh.provenance['psi_policy'])
    evidence = dict(before=before, after=after, accepted=accepted,
        method='one_lsmr_free_original_y_correction_z_unchanged',
        equation='E.T * dy = -(cost + E.T*y - G.T*z)',
        seconds=linear_seconds, maximum_seconds=10., maximum_iterations=200,
        stop=int(correction[1]), iterations=int(correction[2]),
        normr=float(correction[3]), normar=float(correction[4]),
        z_bitwise_unchanged=True, optimizer_calls=0, no_second_attempt=True,
        scope='original dual guide only; no primal bound from linear correction')
    driver.write(args.output / 'linear_repair_audit.json', evidence)
    warm = driver.Certificate('upper', candidate['fs'], mesh, candidate['field'],
        candidate['audit'], str(args.source), False)
    best, root = warm, None
    if accepted:
        guide = dict(case=args.case, mode='upper', q=8, fs=fs, mesh_id=mesh.mesh_id,
            physical_model_id=mesh.physical_model_id, policy=mesh.provenance['psi_policy'],
            dual=repaired)
        # A missing cache entry must fail closed; this run permits no new solve.
        def forbidden_solve(*unused, **ignored):
            raise RuntimeError('NO_OPTIMIZATION_ALLOWED_IN_LINEAR_REPAIR')
        driver.solve = forbidden_solve
        best, root = driver.search(mesh, 'upper', [fs, warm.fs], args.output / 'root',
            args.case + '_linear_repair', 1e-4, 1., initial_witness=warm,
            initial_dual_guides=[guide], budget=driver.CaseBudget(30., 1.), bisection=True)
    current = driver.source_hashes()
    current[str(Path(__file__).resolve())] = sha(__file__)
    unchanged = current == sources and report.unchanged(tracked)
    if not unchanged:
        raise ValueError('Frozen source or input changed')
    metadata = best.publish(args.output)
    result = dict(case=args.case, mode='upper', q=8, fs=best.fs,
        status='OWN_ROOT_COMPLETED' if best.root_complete else 'PHYSICAL_FIELD_ROOT_INCOMPLETE',
        execution_scope=dict(kind='current_session_side_root_continuation',
            historical_field_imported=False, historical_fs_imported=False, optimizer_calls=0),
        source=str(args.source), source_dual=str(args.dual),
        input_sha256={str(p): value for p, value in tracked.items()},
        source_sha256_before=sources, source_sha256_after=current, source_unchanged=unchanged,
        root=root, root_complete=best.root_complete, target_met=False,
        upper=metadata, certificates=dict(upper=metadata), linear_repair=evidence,
        seconds=time.perf_counter() - started)
    driver.write(args.output / 'summary.json', result)
    print(json.dumps(dict(status=result['status'], accepted=accepted,
        linear_seconds=linear_seconds, before=before, after=after,
        seconds=result['seconds'])), flush=True)
    return 0 if best.root_complete else 2


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--case', required=True)
    parser.add_argument('--source', type=Path, required=True)
    parser.add_argument('--dual', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    raise SystemExit(run(parser.parse_args()))
