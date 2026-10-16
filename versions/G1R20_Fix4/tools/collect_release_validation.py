"""Bind completed positive runs to the exact 63 standard sources and event code.

This checks recorded execution provenance; it never runs Excel or a solver.
The raw records remain unchanged, including superseded and failing attempts.
"""
from pathlib import Path
import argparse
import hashlib
import json
from datetime import datetime, timezone

ROOT = Path(__file__).resolve().parents[1]


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def strings(value):
    if isinstance(value, str):
        yield value
    elif isinstance(value, dict):
        for child in value.values():
            yield from strings(child)
    elif isinstance(value, list):
        for child in value:
            yield from strings(child)


def verify_execution_record(record, label):
    """Reject late cleanup failures even in older records that retained PASS.

    Expected input-rejection messages inside cases are test results, so inspect
    only the execution record and its session/stage/run lifecycle records.
    """
    assert record.get('status', record.get('state')) == 'PASS', str(label)
    sessions = record.get('sessions', record.get('stages', []))
    assert record.get('owned_excel_exited') or (sessions and all(s.get('owned_excel_exited') for s in sessions)), str(label)
    immutable = ('monitored_files_unchanged', 'sources_unchanged', 'source_unchanged',
                 'baseline_unchanged', 'source_book_unchanged', 'source_files_unchanged',
                 'original_workbook_unchanged', 'original_workbooks_unchanged')
    failures = ('error', 'verification_error', 'post_run_verification_error',
                'dialogs', 'monitor_errors', 'window_monitor_errors', 'hard_timeout',
                'watchdog_timeout', 'hard_terminated_elapsed', 'watchdog',
                'close_error', 'quit_error', 'book_cleanup_error', 'app_cleanup_error')

    def inspect(item, location):
        for key in immutable:
            if key in item:
                value = item[key]
                valid = all(value.values()) if isinstance(value, dict) else value is True
                assert valid, (str(label), location, key)
        for key in failures:
            assert not item.get(key), (str(label), location, key)
        assert not item.get('published_files_modified'), (str(label), location, 'published_files_modified')
        if 'owned_excel_exited' in item:
            assert item['owned_excel_exited'] is True, (str(label), location, 'owned_excel_exited')
        if 'status' in item or 'state' in item:
            assert item.get('status', item.get('state')) == 'PASS', (str(label), location, 'status')
        if 'exit_code' in item:
            assert item['exit_code'] == 0, (str(label), location, 'exit_code')
        for group in ('sessions', 'stages', 'runs'):
            for index, child in enumerate(item.get(group, [])):
                inspect(child, location + '/' + group + '/' + str(index))

    inspect(record, 'record')


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run', type=Path, action='append', required=True)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if args.output.exists():
        raise FileExistsError('Refusing to replace a validation record')
    manifest = json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    sources = [ROOT/'src/vba'/name for name in manifest['modules']]
    sources += [ROOT/row['utf8_path'] for row in manifest['workbook_modules'].values()]
    expected = {sha(path) for path in sources}
    assert len(sources) == len(expected) == 64
    rows = []
    for run in args.run:
        run = run.resolve()
        assert run.is_relative_to(ROOT/'results')
        path = run/'summary.json'
        if not path.exists():
            path = run/'delivery/WORKBOOK_VALIDATION.json'
        record = json.loads(path.read_text(encoding='utf8'))
        verify_execution_record(record, path)
        found = expected.intersection(strings(record))
        assert found == expected, (str(path), 'Execution record does not bind every final source', len(found))
        rows.append({'record':path.relative_to(ROOT).as_posix(), 'sha256':sha(path),
                     'status':'PASS', 'all_64_source_hashes_bound':True,
                     'owned_excel_exited':True})
    output = {'status':'PASS', 'observed_utc':datetime.now(timezone.utc).isoformat(),
              'scope':'Completed positive execution records and final source provenance; no additional numerical solve',
              'source_bundle_sha256':manifest['source_bundle_sha256'],
              'source_manifest_sha256':sha(ROOT/'SOURCE_MANIFEST.json'),
              'product_sources':{p.relative_to(ROOT).as_posix():sha(p) for p in sources},
              'verifier_sha256':sha(__file__), 'runs':rows}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    args.output.write_text(json.dumps(output, indent=2)+'\n', encoding='utf8')
    print(json.dumps({'status':'PASS', 'positive_runs':len(rows), 'bound_sources':len(sources)}))


if __name__ == '__main__':
    main()
