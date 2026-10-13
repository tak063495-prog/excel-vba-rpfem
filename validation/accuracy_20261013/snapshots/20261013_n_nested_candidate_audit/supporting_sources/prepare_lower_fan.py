"""Prepare a nested lower fan witness only; no solve and no inherited root."""
import argparse
import hashlib
import json
from pathlib import Path
import sys

sys.path.insert(0, r'D:\Work\excel-vba-rpfem\validation\accuracy_20261013\python')
import accuracy_driver as driver
import experiment_core as core
from mesh_accuracy import angular_fan_refine
from upper_directional_candidate import audit_nested_geometry


def run(args):
    _, record = driver.load_model(args.case)
    if args.output.exists():
        raise ValueError('Candidate output must be a new directory')
    mesh = driver.load_saved_mesh(args.seed/'best_lower_mesh.npz', record)
    metadata = json.loads((args.seed/'best_lower_certificate.json').read_text(encoding='utf8'))
    with core.np.load(args.seed/'best_lower_field.npz', allow_pickle=False) as data:
        if str(data['mode']) != 'lower' or int(data['q']) != 8:
            raise ValueError('Source must be a lower q=8 field')
        field, fs = data['x'].copy(), float(data['fs'])
    expected = {'case': args.case, 'mode': 'lower', 'q': 8, 'fs': fs,
                'mesh_id': mesh.mesh_id, 'physical_model_id': mesh.physical_model_id}
    if any(metadata.get(key) != value for key, value in expected.items()):
        raise ValueError('Source field/certificate/mesh/model identity mismatch')
    source = core.evaluate(core.Prepared(mesh, fs, 'lower', 8, reduction='none'), field)
    if not source['physical_pass'] or not source['witness']:
        raise ValueError('Source fails original physical/raw lower gates')
    child, change = angular_fan_refine(mesh, record, scope='star')
    change.update(root_complete=False, geometry_audit=audit_nested_geometry(mesh, child))
    prolonged = core.prolong(mesh, child, field, 'lower', 8)
    result = core.evaluate(core.Prepared(child, fs, 'lower', 8, reduction='none'), prolonged)
    result.update(case=args.case, mode='lower', q=8, fs=fs, mesh_id=child.mesh_id,
                  physical_model_id=child.physical_model_id)
    args.output.mkdir(parents=True)
    driver.save_mesh(args.output/'candidate_mesh.npz', child)
    core.np.savez_compressed(args.output/'candidate_field.npz', x=prolonged, fs=fs, mode='lower', q=8)
    core.write(args.output/'source_audit.json', source)
    core.write(args.output/'candidate_audit.json', result)
    change['source_sha256'] = dict(driver.SOURCE_HASHES)
    for path in [Path(__file__), Path(angular_fan_refine.__code__.co_filename),
                 Path(audit_nested_geometry.__code__.co_filename)]:
        change['source_sha256'][str(path.resolve())] = hashlib.sha256(path.read_bytes()).hexdigest()
    core.write(args.output/'mesh_change.json', change)
    if result['physical_pass'] and result['witness']:
        cert = driver.Certificate('lower', fs, child, prolonged, result,
                                  'audited nested additional angular fan', False)
        core.write(args.output/'import_audit.json', {'case': args.case,
                   'certificates': {'lower': cert.publish(args.output)}, 'mesh_change': change})
    print('LOWER_FAN_CANDIDATE', args.case, len(mesh.tri), len(child.tri),
          result['physical_pass'], result['witness'], result['raw_lower_errors'], flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--case', required=True)
    parser.add_argument('--seed', type=Path, required=True)
    parser.add_argument('--output', type=Path, required=True)
    run(parser.parse_args())
