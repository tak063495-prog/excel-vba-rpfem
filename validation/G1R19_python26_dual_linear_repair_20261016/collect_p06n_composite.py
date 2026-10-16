"""Posthoc original-q8 collection; no solver and no combined-start fiction."""
from pathlib import Path
from datetime import datetime, timezone
import hashlib
import json
import sys

ROOT = Path(__file__).resolve().parent
VALIDATION = ROOT.parent
PIPELINE = VALIDATION / 'G1R19_python26_terminal_final_20261016'
sys.path.insert(0, str(PIPELINE / 'python'))
import accuracy_driver as driver
import report_accuracy as report


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def read(path):
    return json.loads(Path(path).read_text(encoding='utf-8'))


def artifacts(folder):
    return {str(p.resolve()): sha(p) for p in sorted(folder.rglob('*'))
        if p.is_file() and p.suffix.lower() in ('.json', '.npz', '.log')}


def component(base, run_name, label, role, adopted):
    run = VALIDATION / base / 'results' / run_name
    owner_path = run / 'OWNERSHIP.json'
    owner = read(owner_path)
    job = owner['jobs'][label]
    folder = Path(job['folder'])
    summary = read(folder / 'summary.json')
    if not owner['source_unchanged']:
        raise ValueError('A component changed its source')
    return dict(role=role, adopted=adopted, run=str(run), folder=str(folder),
        started_epoch=job['started_epoch'], finished_epoch=job['finished_epoch'],
        process_wall_seconds=job['finished_epoch'] - job['started_epoch'],
        wrapper_seconds=summary.get('seconds'), os_returncode=job['returncode'],
        audit_returncode=job.get('audit_returncode'),
        component_start_source_sha256=owner['source_sha256'],
        component_end_source_sha256=owner['source_sha256_after'], source_unchanged=True,
        ownership_file=dict(path=str(owner_path), sha256=sha(owner_path)),
        summary=dict(path=str(folder / 'summary.json'), sha256=sha(folder / 'summary.json')),
        input_provenance=summary.get('input_sha256'),
        geometry_path=summary.get('geometry_path'), geometry_sha256=summary.get('geometry_sha256'),
        hint=summary.get('hint'), source_path=summary.get('source'),
        execution_scope=summary.get('execution_scope'),
        component_artifacts_sha256=artifacts(folder))


def main():
    stamp = datetime.now(timezone.utc).strftime('%Y%m%dT%H%M%SZ')
    output = ROOT / 'results' / ('P06N_COMPOSITE_' + stamp)
    if output.exists():
        raise ValueError('Fresh output required')
    output.mkdir(parents=True)
    lower_folder = (VALIDATION / 'G1R19_python26_geometry_replay_20261016/results/'
        'known_lower_20261016T122636Z/P06N_known_lower5966')
    upper_folder = ROOT / 'results/P06N_repair_20261016T125042Z'
    _, rec = driver.load_model('P06N')
    candidates = {mode: report.collect_candidate(folder / ('best_' + mode + '_certificate.json'),
        mode, 'P06N', rec, [folder.resolve()], 1e-4)
        for mode, folder in [('lower', lower_folder), ('upper', upper_folder)]}
    if any(not c or not c['root_complete'] for c in candidates.values()):
        raise ValueError('Both original physical witnesses and completed own roots are required')
    if candidates['lower']['mesh'].physical_model_id != candidates['upper']['mesh'].physical_model_id:
        raise ValueError('Physical model differs')
    lower, upper = candidates['lower']['fs'], candidates['upper']['fs']
    gap = 2 * (upper - lower) / (upper + lower)
    if not 0 <= gap <= .02:
        raise ValueError('Unchanged symmetric two-percent gate rejected the pair')
    metadata = {mode: report.freeze_candidate(candidate, output)
        for mode, candidate in candidates.items()}
    specs = [
        ('G1R19_python26_geometry_replay_20261016', 'known_lower_20261016T122636Z',
         'P06N_known_lower5966', 'lower_adopted_fresh_solve', True),
        ('G1R19_python26_upper_goal_recipe_20261016', 'upper_goal_20261016T123423Z',
         'P06N_known_upper8739_goal', 'upper_initial_fresh_lp_qp', True),
        ('G1R19_python26_terminal_final_20261016', 'final_guide_20261016T124646Z',
         'P06N_upper8739_final_guide', 'upper_original_dual_source_before_linear_repair', True),
        ('G1R19_python26_dual_recenter_20261016', 'recenter_20261016T124407Z',
         'P06N_upper8739_recenter_close', 'rejected_original_dual_recenter_reference', False),
    ]
    components = [component(*spec) for spec in specs]
    repair_summary = read(upper_folder / 'summary.json')
    ledger_path = output / 'UNSUPERVISED_REPAIR_EXECUTION_LEDGER.json'
    ledger = dict(kind='captured_exec_without_pid_record', posthoc_record=True,
        command=['D:/Work/rpfem-venv/Scripts/python.exe', '-X', 'utf8',
            'validation/G1R19_python26_dual_linear_repair_20261016/repair_original_dual.py',
            '--case', 'P06N', '--source',
            'validation/G1R19_python26_terminal_final_20261016/results/P06N_immutable_20261016T124645Z',
            '--dual', 'validation/G1R19_python26_terminal_final_20261016/results/'
            'final_guide_20261016T124646Z/P06N_upper8739_final_guide/terminal_original_dual.npz',
            '--output', 'validation/G1R19_python26_dual_linear_repair_20261016/results/'
            'P06N_repair_20261016T125042Z'],
        observed_returncode=0, os_pid=None, os_birth_time=None,
        started_epoch=None, finished_epoch=None, process_wall_seconds=None,
        wrapper_seconds=repair_summary['seconds'], optimizer_calls=0,
        source_summary=dict(path=str(upper_folder / 'summary.json'),
            sha256=sha(upper_folder / 'summary.json')),
        note='Exit0 captured by exec/session49014; PID/birth/start/end OS timing were not recorded. '
             'No cold-process or supervised-exit claim. Source maps below are the actual recorded summary maps.',
        component_start_source_sha256=repair_summary['source_sha256_before'],
        component_end_source_sha256=repair_summary['source_sha256_after'],
        input_sha256=repair_summary['input_sha256'])
    driver.write(ledger_path, ledger)
    repair_component = dict(role='upper_adopted_linear_repair_closure', adopted=True,
        run=str(upper_folder.parent), folder=str(upper_folder), ownership_file=None,
        execution_record=dict(path=str(ledger_path), sha256=sha(ledger_path), **ledger),
        started_epoch=None, finished_epoch=None, process_wall_seconds=None,
        wrapper_seconds=repair_summary['seconds'], os_returncode=0, audit_returncode=0,
        component_start_source_sha256=repair_summary['source_sha256_before'],
        component_end_source_sha256=repair_summary['source_sha256_after'],
        source_unchanged=repair_summary['source_unchanged'],
        summary=dict(path=str(upper_folder / 'summary.json'), sha256=sha(upper_folder / 'summary.json')),
        input_provenance=repair_summary['input_sha256'], source_path=repair_summary['source'],
        component_artifacts_sha256=artifacts(upper_folder),
        independent_primal_audit=dict(path=str(upper_folder) + '_INDEPENDENT_AUDIT.json',
            sha256=sha(str(upper_folder) + '_INDEPENDENT_AUDIT.json')),
        independent_dual_audit=dict(path=str(upper_folder) + '_INDEPENDENT_DUAL_AUDIT.json',
            sha256=sha(str(upper_folder) + '_INDEPENDENT_DUAL_AUDIT.json')))
    components.append(repair_component)
    prior_path = (VALIDATION / 'G1R19_python26_v5_20261016/results/'
        'continuation_20261016T115858Z/TIMING_SUPPLEMENT.json')
    prior = read(prior_path)['cases']['P06N']
    known_adopted = sum(c['process_wall_seconds'] or 0 for c in components if c['adopted'])
    known_all = sum(c['process_wall_seconds'] or 0 for c in components)
    wrapper = repair_summary['seconds']
    manifest = dict(kind='known_recipe_geometry_replay', posthoc_composite_manifest=True,
        created_utc=datetime.now(timezone.utc).isoformat(), not_a_combined_solver_start_manifest=True,
        description='Actual per-component source records; unsupervised repair timing/PID remain explicitly unknown',
        components=components, collector_reaudit_source_sha256=driver.source_hashes(),
        collector_script=dict(path=str(Path(__file__).resolve()), sha256=sha(__file__)),
        rejected_offline_projection_reference=dict(
            folder=str(VALIDATION / 'G1R19_python26_upper_goal_recipe_20261016/results/'
                'offline_dual_projection_20261016T124228Z'),
            adopted=False, process_wall_seconds=None,
            note='No process-wall ledger captured; not silently included as zero measured elapsed'),
        timing=dict(adopted_geometry_solve_component_process_sum_seconds=known_adopted,
            all_recipe_trial_process_sum_seconds=known_all,
            unsupervised_repair_wrapper_seconds=wrapper, unsupervised_repair_process_wall_seconds=None,
            fresh_hint_dependency_v1_plus_v5_seconds=prior['adopted_path_v1_plus_v5_seconds'],
            adopted_components_plus_fresh_hint_path_process_sum_seconds=
                known_adopted + prior['adopted_path_v1_plus_v5_seconds'],
            all_current_session_attempt_process_sum_seconds=
                known_all + prior['all_session_trial_seconds_including_duplicate'],
            prior_timing_file=dict(path=str(prior_path), sha256=sha(prior_path)),
            scope='Process sums contain only supervised measured process durations. Add the separately '
                'measured repair wrapper duration for known elapsed work, not an invented OS duration. '
                'Parallel wall is separate; original archived geometry generation, collector/audits '
                'and the unmeasured offline projection are excluded.'))
    driver.write(output / 'COMPONENT_MANIFEST.json', manifest)
    summary = dict(case='P06N', status='VERIFIED_TARGET', target_met=True, target_gap=.02,
        root_complete=True, lower=metadata['lower'], upper=metadata['upper'],
        relative_width_symmetric=gap,
        execution_scope=dict(kind='known_recipe_geometry_replay', historical_geometry_imported=True,
            historical_field_imported=False, historical_fs_imported=False, generic_cold_start=False,
            posthoc_independent_original_q8_collection=True),
        component_manifest=dict(path='COMPONENT_MANIFEST.json', sha256=sha(output / 'COMPONENT_MANIFEST.json')),
        physics_and_root_gates_unchanged=True, physical_gate_tolerance=1e-7,
        dual_cone_tolerance=1e-8, root_normalized_fs_tolerance=1e-4,
        strict_interval_arithmetic_proof=False)
    driver.write(output / 'summary.json', summary)
    print(json.dumps(dict(output=str(output), lower=lower, upper=upper, gap=gap,
        summary_sha256=sha(output / 'summary.json'),
        component_manifest_sha256=sha(output / 'COMPONENT_MANIFEST.json'),
        timing=manifest['timing'])), flush=True)


if __name__ == '__main__':
    main()
