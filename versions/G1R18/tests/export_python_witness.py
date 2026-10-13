"""Export independently re-audited Python primal fields; never export roots/status."""
from pathlib import Path
import argparse
import hashlib
import json
import struct
import sys

REPO = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO / 'validation/accuracy_20261013/python'))
import accuracy_driver as driver
import experiment_core as core
from import_native import read_state
from experiment_core import np


def export(folder, case, mode, output):
    _, rec = driver.load_model(case)
    paths = [folder / f'best_{mode}_{suffix}' for suffix in ('mesh.npz', 'field.npz', 'certificate.json')]
    metadata = json.loads(paths[2].read_text(encoding='utf-8'))
    mesh = driver.load_saved_mesh(paths[0], rec)
    with np.load(paths[1], allow_pickle=False) as saved:
        field = saved['x'].copy()
        fs, q, saved_mode = float(saved['fs']), int(saved['q']), str(saved['mode'])
    expected = dict(case=case, mode=mode, fs=fs, q=q, mesh_id=mesh.mesh_id,
                    physical_model_id=mesh.physical_model_id, elements=len(mesh.tri))
    if q != 8 or saved_mode != mode or any(metadata.get(k) != v for k, v in expected.items()):
        raise ValueError('Source field/mesh/case/Fs/q metadata disagree')
    if not driver.finite(fs, mesh.stress_scale) or fs <= 0 or not np.isfinite(field).all():
        raise ValueError('Source field must be finite with positive Fs')
    result = core.evaluate(core.Prepared(mesh, fs, mode, 8, reduction='none'), field)
    if not result['witness'] or not result['physical_pass']:
        raise ValueError('Fresh original Python primal/physical gate rejected source')
    pieces = [b'RP25M001', struct.pack('<5i', len(mesh.xy), len(mesh.tri), len(mesh.materials), 0, 8),
              struct.pack('<5d', mesh.stress_scale, fs, fs, result['objective'], 0.),
              np.asarray(mesh.xy, '<f8').tobytes()]
    cells = np.column_stack((mesh.tri, mesh.material_id, mesh.rigid_id)).astype('<i4')
    pieces += [cells.tobytes(), np.asarray(mesh.materials, '<f8').tobytes(),
               np.asarray(mesh.body, '<f8').tobytes(), struct.pack('<i', len(mesh.boundary))]
    for (a, b), boundary in mesh.boundary.items():
        kind = ['load', 'fixed', 'roller_x', 'roller_y'].index(boundary['kind'])
        pieces += [struct.pack('<4i', a, b, kind, boundary['rigid']),
                   np.asarray(boundary['loads'].T, '<f8').tobytes()]
    pieces += [struct.pack('<i', len(field)), np.asarray(field, '<f8').tobytes()]
    raw = b''.join(pieces)
    xy, loaded_cells, mats, body, boundary, loaded_field, loaded_meta = read_state(raw)
    for a, b in ((xy, mesh.xy), (loaded_cells, cells), (mats, mesh.materials),
                 (body, mesh.body), (loaded_field, field)):
        if not np.array_equal(a, b):
            raise ValueError('Binary roundtrip changed original model or field')
    if boundary.keys() != mesh.boundary.keys() or loaded_meta['fs'] != fs:
        raise ValueError('Binary roundtrip changed original boundary/Fs')
    for edge, value in boundary.items():
        original = mesh.boundary[edge]
        if any(value[k] != original[k] for k in ('kind', 'rigid')) or not np.array_equal(value['loads'], original['loads']):
            raise ValueError('Binary roundtrip changed boundary loads')
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_bytes(raw)
    audit = dict(**expected, output=str(output), binary_sha256=hashlib.sha256(raw).hexdigest(),
        original_field_vector_sha256=hashlib.sha256(np.asarray(field,'<f8').tobytes()).hexdigest(),
        ignored_source_metadata={'root_complete': metadata.get('root_complete'),
                                 'solver_status': metadata.get('audit',{}).get('status'),
                                 'imported': False},
        source_sha256={str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in paths},
        source_code_sha256={str(p): hashlib.sha256(p.read_bytes()).hexdigest()
                            for p in (Path(__file__), Path(core.__file__), Path(driver.__file__))},
        fresh_original_audit=result,
        scope='primal field only; no solver status, dual, root, bracket, or precision result imported')
    output.with_suffix('.json').write_text(json.dumps(audit, indent=2) + '\n', encoding='utf-8')
    return audit


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--folder', type=Path, required=True)
    parser.add_argument('--case', required=True)
    parser.add_argument('--mode', choices=['lower', 'upper'], required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    result = export(args.folder.resolve(), args.case, args.mode, args.output.resolve())
    print(json.dumps({k: result[k] for k in ('case','mode','fs','elements','output','binary_sha256')}), flush=True)
