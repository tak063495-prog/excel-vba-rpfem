"""Close an own-mesh root from fresh original primal and original dual audits.

No solver is run: both endpoints must already fit the required root tolerance.
The original dual is a discrete-root guide and never a continuum bound.
"""
import argparse
import hashlib
import json
from pathlib import Path
import shutil
import sys

sys.path.insert(0, r'D:\Work\excel-vba-rpfem\validation\accuracy_20261013\python')
import accuracy_driver as driver
import experiment_core as core


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run(args):
    if args.output.exists():
        raise ValueError('A new independent closure output is required')
    code = Path(__file__).read_bytes()
    implementation = {name: (driver.REPO / name).read_bytes()
                      for name in driver.SOURCE_HASHES}
    implementation_hashes = {name: hashlib.sha256(value).hexdigest()
                             for name, value in implementation.items()}
    if implementation_hashes != driver.SOURCE_HASHES:
        raise ValueError('Implementation changed during import')
    paths = [args.seed / f'best_{args.mode}_{suffix}' for suffix in
             ['certificate.json', 'mesh.npz', 'field.npz']]
    paths += [args.dual / name for name in
              ['audit.json', 'original_mesh.npz', 'original_dual_candidate.npz']]
    inputs = {str(path): sha(path) for path in paths}
    _, record = driver.load_model(args.case)
    metadata = json.loads(paths[0].read_text(encoding='utf8'))
    mesh = driver.load_saved_mesh(paths[1], record)
    with core.np.load(paths[2], allow_pickle=False) as data:
        if (data['q'].shape != () or data['q'].item() != 8
                or data['mode'].shape != () or str(data['mode']) != args.mode
                or data['fs'].shape != ()):
            raise ValueError('A scalar original q8 source certificate is required')
        field, fs = data['x'].copy(), float(data['fs'])
    expected = dict(case=args.case, mode=args.mode, q=8, fs=fs,
                    mesh_id=mesh.mesh_id, physical_model_id=mesh.physical_model_id)
    if not driver.finite(fs) or fs <= 0 or any(metadata.get(k) != v for k, v in expected.items()):
        raise ValueError('Original primal identity mismatch')
    audit = core.evaluate(core.Prepared(mesh, fs, args.mode, 8, reduction='none'), field)
    if not audit['physical_pass'] or not audit['witness']:
        raise ValueError('The full original physical primal audit rejected the source')
    dual_metadata = json.loads(paths[3].read_text(encoding='utf8'))
    dual_mesh = driver.load_saved_mesh(paths[4], record)
    if dual_mesh.mesh_id != mesh.mesh_id or dual_mesh.physical_model_id != mesh.physical_model_id:
        raise ValueError('Original dual mesh differs from the primal mesh')
    if not dual_metadata.get('source_and_program_freeze_pass', False):
        raise ValueError('The original dual source failed its code and program freeze guard')
    with core.np.load(paths[5], allow_pickle=False) as data:
        if any(data[key].shape != () for key in
               ['q', 'fs', 'mode', 'case', 'mesh_id', 'physical_model_id']):
            raise ValueError('A scalar original dual identity is required')
        dual_fs = float(data['fs'])
        dual_expected = dict(expected, fs=dual_fs)
        for key, value in dual_expected.items():
            stored = data[key].item()
            if stored != value or dual_metadata.get(key) != value:
                raise ValueError('Original dual field identity mismatch: ' + key)
        dual = data['dual'].copy()
    lo, hi = (fs, dual_fs) if args.mode == 'lower' else (dual_fs, fs)
    if not driver.finite(lo, hi) or not 0 < lo < hi or hi - lo > 1e-4 * max(1., (lo + hi) / 2):
        raise ValueError('The signed original endpoints do not fit the root tolerance')
    guide = dict(dual_expected, dual=dual, policy=mesh.provenance['psi_policy'])
    # Reassemble at the dual's own strength and audit its unchanged full vector.
    driver.audit_initial_dual_guide(mesh, args.mode, guide)
    audit.update(expected)
    warm = driver.Certificate(args.mode, fs, mesh, field, audit,
                              'fresh full original primal audit', False)
    args.output.mkdir(parents=True)
    for label, base, selected in [('primal_source', args.seed, paths[:3]),
                                  ('dual_source', args.dual, paths[3:])]:
        snapshot = args.output / label
        snapshot.mkdir()
        for path in selected:
            shutil.copyfile(path, snapshot / path.name)
            if sha(snapshot / path.name) != inputs[str(path)]:
                raise ValueError('Frozen input snapshot differs')
    (args.output / 'experiment_source.py').write_bytes(code)
    best, root = driver.search(mesh, args.mode, (lo, hi), args.output,
        args.case + '_' + args.mode + '_original_pair_close', 1e-4, 120,
        initial_witness=warm, initial_dual_guides=[guide])
    frozen = (code == Path(__file__).read_bytes()
        and inputs == {str(path): sha(path) for path in paths}
        and all(value == (driver.REPO / name).read_bytes()
                for name, value in implementation.items()))
    if not frozen:
        best.root_complete = False
        best.result.pop('discrete_root_search', None)
        best.publish(args.output)
        core.write(args.output / 'rejection.json', {'reason': 'FROZEN_SOURCE_CHANGED',
                   'root_complete': False, 'endpoint_promoted': False})
        raise ValueError('Frozen source changed; root closure rejected')
    if not best.root_complete or not root['root_complete'] or best.fs != fs:
        raise ValueError('The unchanged root gate did not complete')
    certificate = best.publish(args.output)
    core.write(args.output / 'import_audit.json', {
        'case': args.case, 'certificates': {args.mode: certificate}})
    core.write(args.output / 'source_provenance.json', {
        'case': args.case, 'mode': args.mode, 'q': 8, 'root': root,
        'original_primal_source': str(args.seed), 'original_dual_source': str(args.dual),
        'source_files_sha256': inputs, 'implementation_sha256': implementation_hashes,
        'candidate_code_sha256': hashlib.sha256(code).hexdigest(),
        'source_and_program_freeze_pass': frozen, 'field_or_root_flags_transferred': False,
        'new_solver_used': False, 'dual_is_continuum_bound': False})
    print('CLOSED_ORIGINAL_PAIR', args.case, args.mode, fs, root['bracket'],
          root['bracket_width'], 'root_complete', best.root_complete, flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--case', required=True)
    parser.add_argument('--mode', choices=['lower', 'upper'], required=True)
    parser.add_argument('--seed', type=Path, required=True)
    parser.add_argument('--dual', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    run(parser.parse_args())
