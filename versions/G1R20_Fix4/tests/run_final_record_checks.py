"""Synthetic final-record checks; no Excel, solver, or raw-record writes.

Run from a full clone to optionally re-audit completed execution records:
python -X utf8 versions/G1R20_Fix4/tests/run_final_record_checks.py --output <new.json> --audit-records
The 24 runner and 12 collector checks are separate from portable regressions.
"""
from pathlib import Path
import argparse
import ast
import copy
import hashlib
import importlib.util
import json

ROOT = Path(__file__).resolve().parents[1]
RUNNERS = ('run_saved_revision_fix4.py', 'run_settings_input_fix4.py',
           'run_seep_state_fix4.py', 'run_settings_extent_display_fix4.py')
GOOD = {'status': 'PASS', 'monitored_files_unchanged': True,
        'sessions': [{'owned_excel_exited': True, 'dialogs': [], 'monitor_errors': []}]}
# The former assertion rejected an error without changing an already saved PASS.
HISTORICAL_CHECK = "assert rec['monitored_files_unchanged'] and all(x['owned_excel_exited'] and not x['dialogs'] and not x['monitor_errors'] and not x.get('hard_timeout') for x in rec['sessions']),rec"


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def final_check(path):
    tree = ast.parse(path.read_text(encoding='utf8'))
    main = next(item for item in tree.body if isinstance(item, ast.FunctionDef) and item.name == 'main')
    statement = main.body[-2]
    assert isinstance(statement, ast.If), (str(path), 'Expected explicit final failure handling')
    return compile(ast.Module(body=[statement], type_ignores=[]), str(path), 'exec')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--audit-records', action='store_true')
    args = parser.parse_args()
    output = args.output.resolve()
    if output.exists():
        raise FileExistsError('Use a new synthetic result file')
    collector_path = ROOT/'tools/collect_release_validation.py'
    spec = importlib.util.spec_from_file_location('final_record_collector', collector_path)
    collector = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(collector)
    report = {'status': 'RUNNING', 'excel_runs': 0, 'solver_calls': 0,
              'scope': 'Synthetic lifecycle checks only; separate 24 runner plus 12 collector cases',
              'historical_check': HISTORICAL_CHECK, 'historical_PASS_retained_reproduced': [],
              'runner_checks': [], 'collector_checks': [], 'positive_raw': [], 'raw_failures': []}
    monitored = [ROOT/'tests'/name for name in RUNNERS] + [collector_path, Path(__file__)]
    report['harness_sha256'] = {p.relative_to(ROOT).as_posix(): sha(p) for p in monitored}
    product = [*sorted((ROOT/'src/vba').glob('*.bas')), ROOT/'src/workbook/ThisWorkbook.code.txt',
               ROOT/'delivery/RPFEM_20261017_G1R20_Fix4.xlsm', ROOT/'delivery/rpfem_native_x64.dll',
               ROOT/'SOURCE_MANIFEST.json']
    product_before = {str(p): sha(p) for p in product}
    for name in RUNNERS:
        path = ROOT/'tests'/name
        rec = copy.deepcopy(GOOD)
        rec['sessions'][0]['monitor_errors'] = ['synthetic']
        saved = []
        try:
            exec(HISTORICAL_CHECK, {'rec': rec, 'save': lambda: saved.append(copy.deepcopy(rec))})
        except AssertionError:
            pass
        else:
            raise AssertionError('Historical check did not reject the monitor error')
        assert rec['status'] == 'PASS' and not saved
        report['historical_PASS_retained_reproduced'].append(name)
        for scenario in ('good', 'immutable', 'exit', 'dialog', 'monitor', 'timeout'):
            rec = copy.deepcopy(GOOD)
            saved = []
            if scenario == 'immutable': rec['monitored_files_unchanged'] = False
            if scenario == 'exit': rec['sessions'][0]['owned_excel_exited'] = False
            if scenario == 'dialog': rec['sessions'][0]['dialogs'] = ['synthetic']
            if scenario == 'monitor': rec['sessions'][0]['monitor_errors'] = ['synthetic']
            if scenario == 'timeout': rec['sessions'][0]['hard_timeout'] = {'seconds': 1}
            rejected = False
            try:
                exec(final_check(path), {'rec': rec, 'save': lambda: saved.append(copy.deepcopy(rec))})
            except AssertionError:
                rejected = True
            assert rejected == (scenario != 'good'), (name, scenario)
            if rejected:
                assert rec['status'] == 'FAIL' and saved[-1]['status'] == 'FAIL', (name, scenario)
            report['runner_checks'].append({'runner': name, 'scenario': scenario, 'passed': True})
    for scenario in ('good', 'root_error', 'nested_monitor', 'nested_dialog', 'timeout', 'immutable',
                     'immutable_map', 'exit', 'child_exit_code', 'child_status', 'published_mutation', 'cleanup_error'):
        rec = copy.deepcopy(GOOD)
        if scenario == 'root_error': rec['post_run_verification_error'] = 'synthetic'
        if scenario == 'nested_monitor': rec['sessions'][0]['monitor_errors'] = ['synthetic']
        if scenario == 'nested_dialog': rec['sessions'][0]['dialogs'] = ['synthetic']
        if scenario == 'timeout': rec['sessions'][0]['hard_timeout'] = {'seconds': 1}
        if scenario == 'immutable': rec['monitored_files_unchanged'] = False
        if scenario == 'immutable_map': rec['source_files_unchanged'] = {'one': False}
        if scenario == 'exit': rec['sessions'][0]['owned_excel_exited'] = False
        if scenario == 'child_exit_code': rec['runs'] = [{'status': 'PASS', 'exit_code': 1}]
        if scenario == 'child_status': rec['runs'] = [{'status': 'FAIL'}]
        if scenario == 'published_mutation': rec['published_files_modified'] = True
        if scenario == 'cleanup_error': rec['sessions'][0]['close_error'] = 'synthetic'
        rejected = False
        try:
            collector.verify_execution_record(rec, 'synthetic '+scenario)
        except AssertionError:
            rejected = True
        assert rejected == (scenario != 'good'), scenario
        report['collector_checks'].append({'scenario': scenario, 'passed': True})
    if args.audit_records:
        historical = dict(zip(RUNNERS, ('saved_revision_positive_01', 'settings_positive_01',
                                       'seep_state_positive_01', 'settings_extent_positive_02')))
        report['historical_harness_snapshots'] = {}
        for name, folder in historical.items():
            snapshot = ROOT/'results'/folder/'runner_snapshot.py'
            if snapshot.exists():
                report['historical_harness_snapshots'][name] = {
                    'path': snapshot.relative_to(ROOT).as_posix(), 'sha256': sha(snapshot),
                    'current_sha256': sha(ROOT/'tests'/name),
                    'current_differs_from_executed_snapshot': sha(snapshot) != sha(ROOT/'tests'/name)}
        # Ignore active runs; never rewrite a completed raw record.
        records = []
        for path in sorted((ROOT/'results').rglob('summary.json')):
            raw = path.read_bytes()
            record = json.loads(raw.decode('utf8'))
            status = record.get('status', record.get('state'))
            if status not in ('PASS', 'FAIL'): continue
            records.append((path, record, hashlib.sha256(raw).hexdigest()))
        for path, record, before in records:
            row = {'path': path.relative_to(ROOT).as_posix(), 'sha256': before}
            if record.get('status', record.get('state')) == 'PASS':
                if record.get('pass_all') and 'legacy_abi' in record: continue
                collector.verify_execution_record(record, path)
                row['no_latecheck_conflict'] = True
                report['positive_raw'].append(row)
            else:
                row['retained_FAIL'] = True
                report['raw_failures'].append(row)
            assert sha(path) == before, ('Raw record changed during read-only audit', str(path))
        report['completed_raw_summaries_unchanged'] = True
    assert all(sha(p) == value for p, value in product_before.items())
    assert all(sha(ROOT/name) == value for name, value in report['harness_sha256'].items())
    manifest = json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    report.update(status='PASS', runner_case_count=24, collector_case_count=12,
                  portable_case_count_increase=0, product_files_unchanged=True,
                  source_bundle_sha256=manifest['source_bundle_sha256'],
                  workbook_sha256=sha(ROOT/'delivery/RPFEM_20261017_G1R20_Fix4.xlsm'))
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, indent=2)+'\n', encoding='utf8')
    print(json.dumps({'status': 'PASS', 'runner_cases': 24, 'collector_cases': 12,
                      'positive_raw': len(report['positive_raw']), 'retained_FAIL': len(report['raw_failures']),
                      'output': str(output)}))


if __name__ == '__main__':
    main()
