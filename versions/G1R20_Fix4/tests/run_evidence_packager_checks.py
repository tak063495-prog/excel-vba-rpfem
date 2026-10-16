"""Eight synthetic actual-main evidence-packager checks; no Excel or solver.

Temporary runs and archives are created only in a fresh scratch directory.
These eight cases are separate from the 132 portable and 24+12 lifecycle cases.
"""
from pathlib import Path
import argparse
import contextlib
import copy
from datetime import datetime, timezone
import hashlib
import importlib.util
import io
import json
import os
import sys
import tempfile
import zipfile
import psutil

ROOT = Path(__file__).resolve().parents[1]
REPO = ROOT.parents[1]


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--scratch-parent', type=Path, default=REPO/'scratch')
    args = parser.parse_args()
    output = args.output.resolve()
    if output.exists():
        raise FileExistsError('Use a new packager-check result file')
    parent = args.scratch_parent.resolve()
    assert parent.is_relative_to(REPO) or parent == REPO.parent, 'Scratch must be in the workspace'
    parent.mkdir(parents=True, exist_ok=True)
    scratch = Path(tempfile.mkdtemp(prefix='fix4-evidence-packager-', dir=parent))
    tool = ROOT/'tools/package_test_evidence.py'
    spec = importlib.util.spec_from_file_location('synthetic_evidence_packager', tool)
    packager = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(packager)
    monitored = [*sorted((ROOT/'src/vba').glob('*.bas')), ROOT/'src/workbook/ThisWorkbook.code.txt',
                 ROOT/'delivery/RPFEM_20261017_G1R20_Fix4.xlsm', ROOT/'delivery/rpfem_native_x64.dll',
                 ROOT/'SOURCE_MANIFEST.json', tool, Path(__file__)]
    before = {str(path): sha(path) for path in monitored}
    # Current product raw records are never used as synthetic fixtures or edited.
    raw = {str(path): sha(path) for path in (ROOT/'results').rglob('summary.json')}
    absent_pid = max(psutil.pids()) + 100000
    assert not psutil.pid_exists(absent_pid)
    birth = psutil.Process(os.getpid()).create_time()
    observed = datetime.now(timezone.utc).isoformat()
    cases = [('preserved_false_post_true', True), ('post_false', False),
             ('pid_mismatch', False), ('birth_mismatch', False), ('live_exact_python_owner', False),
             ('process_present_true', False), ('legacy_single_observation', True), ('ambiguous_observation', False)]
    report = {'status': 'RUNNING', 'scope': 'Actual packager main with synthetic scratch-only records',
              'excel_runs': 0, 'solver_calls': 0, 'processes_terminated': 0, 'case_count': 8,
              'portable_case_count_increase': 0, 'lifecycle_case_count_increase': 0,
              'scratch': str(scratch), 'packager_sha256': sha(tool), 'runner_sha256': sha(__file__), 'cases': []}
    for name, expected in cases:
        case_repo = scratch/name/'repo'
        case_root = case_repo/'versions/G1R20_Fix4'
        run = case_root/'results/synthetic_run'
        run.mkdir(parents=True)
        pid = os.getpid() if name == 'live_exact_python_owner' else absent_pid
        summary = {'status': 'FAIL', 'owned_excel_pid': pid, 'owned_excel_birth': birth,
                   'owned_excel_exited': False, 'error': 'Synthetic original teardown failure'}
        observation = {'pid': pid, 'creation_time': birth, 'original_owner_exited': False,
                       'original_owned_excel_exited': False, 'post_run_owned_excel_exited': True,
                       'owned_excel_exited': True, 'process_present': False}
        later = {'observed_utc': observed, 'sessions': [observation]}
        if name == 'post_false':
            observation.update(post_run_owned_excel_exited=False, original_owner_exited=True)
        if name == 'pid_mismatch': observation['pid'] += 1
        if name == 'birth_mismatch': observation['creation_time'] += 1
        if name == 'process_present_true': observation['process_present'] = True
        if name == 'legacy_single_observation':
            later = {'observed_utc': observed, 'pid': pid, 'creation_time': birth,
                     'original_owner_exited': True, 'ownership_exit_verified': True}
        if name == 'ambiguous_observation': later['sessions'].append(copy.deepcopy(observation))
        source = run/'summary.json'
        source.write_text(json.dumps(summary, indent=2)+'\n', encoding='utf8')
        post = run/'POST_RUN_EXIT_CHECK.json'
        post.write_text(json.dumps(later, indent=2)+'\n', encoding='utf8')
        original = {str(path): sha(path) for path in (source, post)}
        archive = scratch/name/'evidence.zip'
        saved_argv = sys.argv
        packager.ROOT, packager.REPO = case_root, case_repo
        accepted = False
        error = None
        try:
            sys.argv = [str(tool), '--output', str(archive), '--run', str(run)]
            with contextlib.redirect_stdout(io.StringIO()):
                packager.main()
            accepted = True
        except (AssertionError, ValueError) as failure:
            error = type(failure).__name__ + ': ' + str(failure)
        finally:
            sys.argv = saved_argv
        assert accepted == expected, (name, accepted, expected, error)
        assert all(sha(path) == value for path, value in original.items()), (name, 'Synthetic raw changed')
        if accepted:
            with zipfile.ZipFile(archive) as zipped:
                assert zipped.testzip() is None
                manifest = json.loads(zipped.read('MANIFEST.json'))
                row = manifest['summaries'][0]
                assert row['status'] == 'FAIL' and row['original_owned_excel_exited'] is False
                assert row['later_exit_observation'] == later
                assert zipped.read('versions/G1R20_Fix4/results/synthetic_run/summary.json') == source.read_bytes()
            assert archive.with_suffix('.json').exists()
        else:
            assert not archive.exists(), (name, 'Rejected observation created an archive')
        report['cases'].append({'case': name, 'expected_acceptance': expected, 'accepted': accepted,
                                'passed': True, 'rejection': error, 'synthetic_raw_unchanged': True})
    assert all(sha(path) == value for path, value in before.items()), 'Product or harness changed'
    assert all(sha(path) == value for path, value in raw.items()), 'Existing raw summary changed'
    report.update(status='PASS', product_and_harness_unchanged=True, existing_raw_summaries_unchanged=True,
                  existing_raw_summary_count=len(raw), product_and_harness_sha256=before,
                  workbook_sha256=sha(ROOT/'delivery/RPFEM_20261017_G1R20_Fix4.xlsm'))
    output.parent.mkdir(parents=True, exist_ok=True)
    output.write_text(json.dumps(report, indent=2)+'\n', encoding='utf8')
    print(json.dumps({'status': 'PASS', 'cases': 8, 'output': str(output), 'scratch': str(scratch)}))


if __name__ == '__main__':
    main()
