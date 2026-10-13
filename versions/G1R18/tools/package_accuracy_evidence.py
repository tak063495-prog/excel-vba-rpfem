"""Curate immutable small Excel evidence; never include running production files."""
from pathlib import Path
import hashlib
import json
import zipfile

ROOT = Path(__file__).resolve().parents[1]
FOLDERS = {
    'adapt_excel_1791896088836241400': 'Angular fan geometry/material/body/boundary/cap integrity PASS',
    'adapt_excel_1791896964526437200': 'Independent own q8 root/field checks and injected-side q retention PASS',
    'adapt_excel_1791897436658252400': 'Retained pre-fix lower layout regression failure',
    'adapt_excel_1791898363593743200': 'Lower layout restore and independent lower uniform/fan improvement PASS',
    'adapt_excel_1791900201404917400': 'Retained nonfinite wedge depth preflight failure before finite/range split',
    'adapt_excel_1791900420176397000': 'Retained test-only root output compound-expression failure',
    'adapt_excel_1791900646125391300': 'Upper wedge adoption/lower exact retention/injected failure rollback PASS',
    'adapt_excel_1791901989652207900': 'Final original Farkas gate/capture/native q8 wedge and independent matrix replay PASS',
}


def main():
    results = ROOT / 'results'
    target = results / 'accuracy_control_evidence_20261013.zip'
    files = []
    runs = {}
    for name, scope in FOLDERS.items():
        folder = results / name
        summary = json.loads((folder / 'summary.json').read_text(encoding='utf-8'))
        assert summary['status'] in ('PASS', 'FAIL'), 'Refusing live production evidence'
        runs[name] = {'status': summary['status'], 'scope': scope}
        selected = [folder / 'summary.json']
        selected += list(folder.glob('original_evidence_replay.json'))
        selected += list(folder.glob('*.negative_*.bin'))
        selected += list(folder.glob('*.negative_*.tsv'))
        selected += [p for p in (folder / 'RPFEM_logs').glob('*') if p.suffix in ('.tsv', '.bin')]
        files.extend(sorted(selected))
    manifest = {'scope': 'Small control-flow and original-model tests; no P06 production success claim',
        'production_outputs_excluded': True, 'workbooks_and_unchanged_dll_excluded': True,
        'runs': runs, 'files': {str(p.relative_to(results)).replace('\\', '/'):
            {'sha256': hashlib.sha256(p.read_bytes()).hexdigest(), 'bytes': p.stat().st_size} for p in files}}
    manifest_bytes = (json.dumps(manifest, indent=2) + '\n').encode('utf-8')
    with zipfile.ZipFile(target, 'w', compression=zipfile.ZIP_DEFLATED, compresslevel=9) as archive:
        archive.writestr('manifest.json', manifest_bytes)
        for path in files:
            archive.write(path, str(path.relative_to(results)).replace('\\', '/'))
    with zipfile.ZipFile(target) as archive:
        assert archive.testzip() is None
        for name, record in manifest['files'].items():
            assert hashlib.sha256(archive.read(name)).hexdigest() == record['sha256']
    (results / 'accuracy_control_evidence_20261013_manifest.json').write_bytes(manifest_bytes)
    print(json.dumps({'archive': str(target), 'archive_bytes': target.stat().st_size,
                      'files': len(files), 'source_bytes': sum(p.stat().st_size for p in files)}))


if __name__ == '__main__':
    main()
