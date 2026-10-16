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
        assert record.get('status', record.get('state')) == 'PASS', str(path)
        sessions = record.get('sessions', record.get('stages', []))
        assert record.get('owned_excel_exited') or (sessions and all(s.get('owned_excel_exited') for s in sessions)), str(path)
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
