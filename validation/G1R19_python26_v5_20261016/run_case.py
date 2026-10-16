"""Bounded generic pilot policy; only identified fresh runs may continue."""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import sys
import time

ROOT = Path(__file__).resolve().parent
sys.path.insert(0, str(ROOT/'python'))
import accuracy_driver as driver


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def snapshot_predecessor(seed, prior, case, destination):
    seed, prior = seed.resolve(), prior.resolve()
    if not seed.is_relative_to(prior):
        raise ValueError('Seed must belong to the explicitly identified fresh predecessor')
    outer = json.loads((prior/'summary.json').read_text(encoding='utf-8'))
    inner = json.loads((seed/'summary.json').read_text(encoding='utf-8'))
    allowed = ('cold_start_new_python_policy', 'fresh_automatic_candidate_then_bounded_generic_policy',
               'fresh_cold_start_with_new_generic_policy', 'fresh_pipeline_continuation')
    if (outer.get('case') != case or inner.get('case') != case or
            outer.get('execution_scope', {}).get('kind') not in allowed or
            inner.get('execution_scope', {}).get('kind') not in allowed):
        raise ValueError('Historical/saved-witness or unbound predecessor is not a new cold-run continuation')
    if inner.get('status') == 'RUNNING' or outer.get('status') == 'RUNNING':
        raise ValueError('Wait for the frozen predecessor to stop before snapshotting')
    names = ['summary.json']+[f'best_{mode}_{suffix}' for mode in ('lower', 'upper')
                for suffix in ('certificate.json', 'mesh.npz', 'field.npz')]
    original = {name: (seed/name).read_bytes() for name in names}
    destination.mkdir(parents=True)
    for name, raw in original.items():
        (destination/name).write_bytes(raw)
    if any((seed/name).read_bytes() != raw or (destination/name).read_bytes() != raw
           for name, raw in original.items()):
        raise ValueError('Predecessor changed during exact input snapshot')
    elapsed = max(float(report.get('case_budget', {}).get('elapsed_seconds', 0.))
                  for report in (outer, inner))
    elapsed = max(elapsed, float(outer.get('cold_cumulative_seconds', 0.)))
    if not driver.finite(elapsed) or elapsed < 0:
        raise ValueError('Invalid declared predecessor duration')
    return dict(prior_run=str(prior), seed=str(seed), elapsed_seconds=elapsed,
        original_status=outer.get('status'), seed_status=inner.get('status'),
        input_sha256={name: hashlib.sha256(raw).hexdigest() for name, raw in original.items()},
        historical_witness_imported=False, predecessor_outer_summary_sha256=sha(prior/'summary.json'))


def parser():
    p = argparse.ArgumentParser()
    p.add_argument('--case', required=True)
    p.add_argument('--output', type=Path, required=True)
    p.add_argument('--seed', type=Path)
    p.add_argument('--prior-run', type=Path, help='Required for a fresh-pipeline seed, including its whole elapsed time')
    p.add_argument('--additional-attempt-report',type=Path,action='append',default=[],
                   help='Timing-only report of an unadopted current-session comparison')
    p.add_argument('--initial-mesh', type=Path)
    p.add_argument('--cycles', type=int, default=8)
    p.add_argument('--max-elements', type=int, default=32000)
    p.add_argument('--theta', type=float, default=.8)
    p.add_argument('--lower-theta', type=float, default=.8)
    p.add_argument('--target-gap', type=float, default=.02)
    p.add_argument('--tolerance', type=float, default=1e-4)
    p.add_argument('--case-seconds', type=float, default=900.)
    p.add_argument('--solve-seconds', type=float, default=30.)
    p.add_argument('--fan-lower', action='store_true')
    p.add_argument('--shared-lower-mesh', action='store_true', help='Comparison only: keep the previous shared lower-mesh strategy')
    p.add_argument('--upper-indicator', choices=['mixed', 'kinematic'], default='mixed')
    p.add_argument('--kinematic-fraction', type=float, default=.2)
    p.add_argument('--initial-span', type=float, default=.005)
    p.add_argument('--root-method', choices=['bisection', 'secant'], default='bisection')
    p.add_argument('--pilot-trials',type=int,default=4)
    p.add_argument('--recheck', action='store_true')
    p.add_argument('--upper-only', action='store_true')
    return p


def run(args):
    started = time.perf_counter()
    original_budget = args.case_seconds
    wrapper_sha = sha(Path(__file__))
    static_inputs=[ROOT/'INPUTS_26.json',driver.CATALOG/(args.case+'.json')]+args.additional_attempt_report
    static_input_sha={str(path):sha(path) for path in static_inputs}
    other_attempts=[]
    for path in args.additional_attempt_report:
        data=json.loads(path.read_text(encoding='utf-8'))
        entry=data.get('cases',{}).get(args.case,data if data.get('case')==args.case else {})
        seconds=entry.get('elapsed_seconds')
        if not driver.finite(seconds) or seconds<0:raise ValueError('Comparison report needs its measured case elapsed time')
        other_attempts.append(dict(path=str(path.resolve()),sha256=sha(path),elapsed_seconds=float(seconds),
            status=entry.get('status',data.get('status')),fields_imported=False,scope='timing only; unadopted comparison'))
    if bool(args.seed) != bool(args.prior_run):
        raise ValueError('A continuation needs both --seed and --prior-run; old saved proofs are not accepted')
    if args.seed and args.initial_mesh:
        raise ValueError('Mesh-only cold start and continuation are separate scopes')
    folder = args.output.resolve()/args.case
    if folder.exists():
        raise ValueError('A fresh output folder is required')
    predecessor = None
    mesh_provenance = None
    if args.seed:
        predecessor = snapshot_predecessor(args.seed, args.prior_run, args.case, folder/'predecessor_input_snapshot')
        args.seed = folder/'predecessor_input_snapshot'
    if args.initial_mesh:
        source_mesh=args.initial_mesh.resolve()
        prefix=source_mesh.stem.removesuffix('_mesh')
        paths=[source_mesh,source_mesh.with_name(prefix+'.context.tsv'),
               source_mesh.with_name(prefix+'.bin'),source_mesh.parent/'input_record.json',
               source_mesh.parent.parent/'summary.json']
        paths=[p for p in paths if p.is_file()]
        before={str(p):p.read_bytes() for p in paths}
        destination=folder/'mesh_only_source_snapshot'
        destination.mkdir(parents=True)
        names={}
        for i,(path,raw) in enumerate(before.items()):
            name=f'{i:02d}_{Path(path).name}'
            copied=destination/name;copied.write_bytes(raw);names[path]=name
            if copied.read_bytes()!=raw:raise ValueError('Mesh-only source copy changed')
        if any(Path(path).read_bytes()!=raw for path,raw in before.items()):
            raise ValueError('Mesh-only source changed during snapshot')
        args.initial_mesh=destination/names[str(source_mesh)]
        mesh_provenance=dict(source=str(source_mesh), imports_fields_or_roots=False,
            input_sha256={path:hashlib.sha256(raw).hexdigest() for path,raw in before.items()},
            snapshot_files=names, generation_time_is_separate_from_python_solver_time=True,
            geometry_generation_seconds=None)
    args.resume = False
    args.independent_lower = not args.shared_lower_mesh
    args.case_seconds -= time.perf_counter()-started
    if args.case_seconds <= 0:
        raise ValueError('Whole-case budget expired before solving')
    report = driver.run(args)
    elapsed = time.perf_counter()-started
    report['execution_scope'].update(kind='fresh_pipeline_continuation' if predecessor else 'fresh_cold_start_with_new_generic_policy',
        description='New generic Python policy; no historical proof or Excel timing claim',
        historical_witness_imported=False, fresh_predecessor=predecessor,
        mesh_only_provenance=mesh_provenance,
        inherited_root_is_mesh_specific=True, field_transport_always_reaudited=True)
    report['cold_cumulative_seconds'] = elapsed+(predecessor['elapsed_seconds'] if predecessor else 0.)
    report['attempt_kind']='continuation' if predecessor else 'fresh_initial_attempt'
    report['adopted_path_seconds']=report['cold_cumulative_seconds']
    report['other_current_session_attempts']=other_attempts
    report['all_current_session_trial_seconds']=report['adopted_path_seconds']+sum(r['elapsed_seconds'] for r in other_attempts)
    report['trial_time_scope']='Per-case sum of measured attempt durations; parallel batch wall time is separate'
    report['cold_cumulative_time_scope']='Sum of fresh Python process intervals; external native mesh generation timing is separate and was not recorded'
    report['pipeline_version']=5
    report['process_elapsed_seconds'] = elapsed
    report['whole_process_budget_seconds'] = original_budget
    report['wrapper_source_sha256'] = wrapper_sha
    report['wrapper_source_unchanged'] = wrapper_sha == sha(Path(__file__))
    report['frozen_input_sha256']=static_input_sha
    report['frozen_inputs_unchanged']=all(sha(Path(path))==digest for path,digest in static_input_sha.items())
    if mesh_provenance:
        report['mesh_only_source_unchanged']=all(sha(Path(p))==digest for p,digest in mesh_provenance['input_sha256'].items())
        if not report['mesh_only_source_unchanged']:
            report['target_met']=False;report['status']='ERROR'
    if not report['wrapper_source_unchanged'] or not report['frozen_inputs_unchanged'] or elapsed > original_budget:
        report['target_met'] = False
        report['status'] = 'ERROR' if not report['wrapper_source_unchanged'] or not report['frozen_inputs_unchanged'] else 'BUDGET_EXHAUSTED'
    driver.write(folder/'summary.json', report)
    print('V5_FINAL', args.case, report['status'], report.get('relative_width_symmetric'),
          'cumulative_seconds', report['cold_cumulative_seconds'], flush=True)
    return report


if __name__ == '__main__':
    result = run(parser().parse_args())
    raise SystemExit(0 if result.get('target_met') else 2)
