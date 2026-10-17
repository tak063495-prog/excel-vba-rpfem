"""Fresh process original-q8 dual replay; no optimization."""
from pathlib import Path
import hashlib
import json
import sys
import time

ROOT = Path(__file__).resolve().parent
PIPELINE = ROOT.parent / 'G1R19_python26_terminal_final_20261016'
sys.path.insert(0, str(PIPELINE / 'python'))
import accuracy_driver as driver
from original_dual import audit_original_dual


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def main():
    started = time.perf_counter()
    folder = ROOT / 'results/P06N_repair_20261016T125042Z'
    output = ROOT / 'results/P06N_repair_20261016T125042Z_INDEPENDENT_DUAL_AUDIT.json'
    if output.exists():
        raise ValueError('Fresh audit output required')
    paths = [folder / 'repaired_original_dual.npz', folder / 'best_upper_mesh.npz',
        driver.CATALOG / 'P06N.json', Path(__file__)]
    before = {str(p.resolve()): sha(p) for p in paths}
    sources = driver.source_hashes()
    _, record = driver.load_model('P06N')
    with driver.np.load(paths[1], allow_pickle=False) as data:
        mesh = driver.core.mesh_from_arrays(data['xy'].copy(), data['tri'].copy(),
            data['material_id'].copy(), record)
    with driver.np.load(paths[0], allow_pickle=False) as data:
        fs, dual = float(data['fs']), data['dual'].copy()
        identity = dict(case=str(data['case']), mode=str(data['mode']), q=int(data['q']),
            mesh_id=str(data['mesh_id']), physical_model_id=str(data['physical_model_id']),
            policy=str(data['policy']))
    expected = dict(case='P06N', mode='upper', q=8, mesh_id=mesh.mesh_id,
        physical_model_id=mesh.physical_model_id, policy=mesh.provenance['psi_policy'])
    if identity != expected:
        raise ValueError('Full vector context differs from reconstructed original model')
    prepared = driver.core.Prepared(mesh, fs, 'upper', 8, reduction='none')
    audit = audit_original_dual(prepared, dual)
    after = {str(p.resolve()): sha(p) for p in paths}
    if before != after or sources != driver.source_hashes():
        raise ValueError('Audit source or input bytes changed')
    result = dict(status='VERIFIED_ORIGINAL_DUAL' if audit['accepted'] else 'REJECTED',
        fs=fs, context=identity, audit=audit, seconds=time.perf_counter() - started,
        optimizer_calls=0, code_and_input_sha256_before=before,
        code_and_input_sha256_after=after, source_sha256=sources, source_unchanged=True,
        evidence_scope='opposite discrete sign only; not a continuum bound')
    driver.write(output, result)
    print(json.dumps(dict(status=result['status'], fs=fs, audit=audit,
        seconds=result['seconds'])), flush=True)
    return 0 if audit['accepted'] else 2


if __name__ == '__main__':
    raise SystemExit(main())
