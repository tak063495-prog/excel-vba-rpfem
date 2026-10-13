"""Refine an independently audited lower witness and preserve its own mesh.

Example: python -X utf8 lower_accuracy_driver.py --case P06A --method uniform
    --seed ../native_seeds/P06A --output ../results/lower_uniform --rounds 1
Each round exports a lower-only import_audit.json for combine_seeds.py and
report_accuracy.py. No upper field or one-percent result is inferred here.
"""
from __future__ import annotations
import argparse
import hashlib
import json
import platform
from pathlib import Path
import accuracy_driver as driver
import experiment_core as core
from mesh_accuracy import angular_fan_refine, lower_stress_score

SOURCE_HASHES = dict(driver.SOURCE_HASHES)
SOURCE_HASHES[str(Path(__file__).resolve().relative_to(driver.REPO))] = hashlib.sha256(Path(__file__).read_bytes()).hexdigest()


def refine_lower(mesh, record, field, fs, method, theta, proximity_background=.1):
    if method == 'uniform':
        child = core.uniform_refine(mesh, record)
        return child, {'method': 'uniform_red', 'elements_before': len(mesh.tri),
                       'elements_after': len(child.tri)}, None
    if method != 'stress_fan':
        raise ValueError('Unknown lower refinement method')
    score, indicator = lower_stress_score(mesh, field, fs, with_metadata=True,
                                          proximity_background=proximity_background)
    refined, local = core.local_refine(mesh, record, score, theta)
    child, fan = angular_fan_refine(refined, record)
    # angular_fan_refine maps to its immediate input. Compose both levels so
    # the original stress polynomial can be evaluated directly in the child.
    ancestry = refined.provenance['parent_element_ids']
    child.provenance['parent_element_ids'] = [ancestry[e] for e in child.provenance['parent_element_ids']]
    return child, {'method': 'stress_variation_then_single_angular_fan',
                   'local': local, 'fan': fan, 'elements_before': len(mesh.tri),
                   'elements_after': len(child.tri)}, indicator


def seed_metadata(folder):
    if (folder/'summary.json').is_file():
        data = json.loads((folder/'summary.json').read_text(encoding='utf8'))
    elif (folder/'import_audit.json').is_file():
        data = json.loads((folder/'import_audit.json').read_text(encoding='utf8'))
    else:
        raise ValueError('The seed needs summary.json or import_audit.json')
    return data.get('certificates', data).get('lower', {})


def publish(folder, case, certificate, history):
    folder.mkdir(parents=True, exist_ok=True)
    lower = certificate.publish(folder)
    core.write(folder/'import_audit.json', {'case': case, 'certificates': {'lower': lower},
               'scope': 'Lower-only original physical witness; no inferred upper bound'})
    core.write(folder/'summary.json', {'case': case, 'lower': lower, 'lower_only': True,
               'one_percent_target_not_evaluated': True, 'history': history,
               'environment': {'python': platform.python_version(), 'numpy': core.np.__version__,
                               'clarabel': core.clarabel.__version__, 'threads': 1,
                               'source_sha256': SOURCE_HASHES}})
    return lower


def run(args):
    if args.rounds < 0 or args.max_elements <= 0 or not driver.finite(args.theta,args.tolerance,args.solve_seconds,args.proximity_background) or not 0 < args.theta <= 1 or not 0 < args.tolerance <= 1e-4 or args.solve_seconds <= 0 or not 0 <= args.proximity_background <= 1:
        raise ValueError('Rounds, element cap, marking fraction, tolerance, or solve budget is invalid')
    _, record = driver.load_model(args.case)
    seed = args.seed.resolve()
    metadata = seed_metadata(seed)
    mesh = driver.load_saved_mesh(seed/'best_lower_mesh.npz', record)
    with core.np.load(seed/'best_lower_field.npz', allow_pickle=False) as data:
        if str(data['mode']) != 'lower' or int(data['q']) != 8:
            raise ValueError('The seed must be a lower witness with q=8')
        field, fs = data['x'].copy(), float(data['fs'])
    if not driver.finite(fs) or fs <= 0 or metadata.get('fs') != fs or metadata.get('mesh_id') != mesh.mesh_id or metadata.get('physical_model_id') != mesh.physical_model_id:
        raise ValueError('The stored seed field, mesh, model, and strength factor differ')
    result = core.evaluate(core.Prepared(mesh, fs, 'lower', 8, reduction='none'), field)
    if not result['physical_pass'] or not result['witness']:
        raise ValueError('The unchanged original physical gate rejected the seed')
    # Prior root flags are evidence only; a new own-mesh search establishes
    # completion in every produced round. A rounds=0 export remains unclosed.
    best = driver.Certificate('lower', fs, mesh, field, result, str(seed), False)
    folder = args.output.resolve()/args.case
    history = []
    publish(folder, args.case, best, history)
    for number in range(1, args.rounds+1):
        child, change, indicator = refine_lower(mesh, record, field, fs, args.method, args.theta,
                                                args.proximity_background)
        if len(child.tri) > args.max_elements:
            history.append({'round': number, 'elements': len(child.tri), 'reason': 'ELEMENT_CAP', 'refinement': change})
            publish(folder, args.case, best, history)
            print('ELEMENT_CAP', args.case, len(child.tri), flush=True)
            break
        label = f'{args.case}_lower_{args.method}_R{number}'
        round_folder = folder/label
        round_folder.mkdir(parents=True, exist_ok=True)
        driver.save_mesh(round_folder/'candidate_mesh.npz', child)
        core.write(round_folder/'mesh_change.json', change)
        if indicator is not None:
            core.write(round_folder/'stress_indicator.json', indicator)
        prolonged = core.prolong(mesh, child, field, 'lower')
        warm_result = core.evaluate(core.Prepared(child, fs, 'lower', 8, reduction='none'), prolonged)
        if not warm_result['physical_pass'] or not warm_result['witness']:
            raise ValueError('The prolonged stress polynomial failed the unchanged physical gate')
        warm_result.update(case=args.case, fs=fs, mode='lower', q=8, mesh_id=child.mesh_id,
                           physical_model_id=child.physical_model_id, search_sign=1)
        core.write(round_folder/'prolonged_audit.json', warm_result)
        warm = driver.Certificate('lower', fs, child, prolonged, warm_result,
                                  'independently audited polynomial prolongation', False)
        publish(round_folder, args.case, warm, [])
        print('WARM_LOWER', args.case, number, len(child.tri), fs, warm_result['value'], flush=True)
        candidate, root = driver.search(child, 'lower', (fs, fs+.035), round_folder, label,
                         args.tolerance, args.solve_seconds, initial_witness=warm)
        candidate = candidate or warm
        row = {'round': number, 'method': args.method, 'elements': len(child.tri), 'refinement': change,
               'fs_lower': candidate.fs, 'root': root}
        publish(round_folder, args.case, candidate, [row])
        history.append(row)
        if candidate.fs > best.fs or (candidate.fs == best.fs and candidate.root_complete and not best.root_complete):
            best = candidate
        publish(folder, args.case, best, history)
        mesh, field, fs = child, candidate.field, candidate.fs
        print('LOWER_BOUND', args.case, number, len(child.tri), fs, candidate.root_complete, flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--case', required=True)
    parser.add_argument('--seed', type=Path, required=True)
    parser.add_argument('--method', choices=['uniform', 'stress_fan'], required=True)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--rounds', type=int, default=1)
    parser.add_argument('--max-elements', type=int, default=24000)
    parser.add_argument('--theta', type=float, default=.6)
    parser.add_argument('--proximity-background', type=float, default=.1,
                        help='Stress-score weight away from yield: zero emphasizes yielded regions; default .1')
    parser.add_argument('--tolerance', type=float, default=1e-4)
    parser.add_argument('--solve-seconds', type=float, default=120)
    run(parser.parse_args())
