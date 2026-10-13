"""Freeze selected completed native audits; never archive a live workbook."""
from pathlib import Path
import argparse
import hashlib
import json
import zipfile

VARIANT = Path(__file__).resolve().parents[1]
BASE = VARIANT.parent / 'G1R18'
DEFAULTS = [
    'python_field_excel_1791903229682201900',
    'python_field_excel_1791903712492988200',
    'python_field_excel_1791903997288551900',
    'python_field_excel_1791904224800777700',
    'python_field_excel_1791905614894344800',
    'adapt_excel_1791904907321458900',  # rejected strict AHO3 linear recovery
    'adapt_excel_1791907938957103600',  # rejected near-root AF HSD
    'adapt_excel_1791908292891767200',  # AF recovery / exact rollback
    'adapt_excel_1791908375927363000',  # passive original ray / wedge / root
    'adapt_excel_1791910801411367300',  # recovered primal / separate free dual root / finite working cost
]


def freeze(names, name):
    if not name or any(character not in 'abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789_-' for character in name):
        raise ValueError('Archive name must be one safe filename component')
    files, runs = {}, []
    for item in names:
        folder = BASE / 'results' / item
        if folder.parent.resolve() != (BASE / 'results').resolve():
            raise ValueError('Only an explicitly named direct completed results directory is allowed')
        summary = json.loads((folder / 'summary.json').read_text(encoding='utf-8'))
        if summary['status'] not in ('PASS', 'FAIL'):
            raise ValueError('Refusing live evidence: ' + item)
        if summary['status'] == 'FAIL' and item not in DEFAULTS:
            raise ValueError('Additional failures require explicit curated source selection')
        runs.append(dict(directory=item, status=summary['status'], pid=summary['pid'],
                         source_sha256=summary.get('source_sha256'),
                         source=summary.get('source'), results=summary.get('runs'),
                         purpose='Completed audit/regression; no native precision achievement inferred'))
        for path in sorted(folder.rglob('*')):
            if not path.is_file() or path.suffix.lower() not in ('.json', '.bin', '.tsv', '.log', '.txt'):
                continue
            if path.name in ('q4.bin', 'nan_xy.bin', 'changed_material.bin'):
                continue  # explicit rejection checks are retained in the completed summary
            if path.name in ('rpfem_native_x64.dll', 'production_metadata_live.json'):
                raise ValueError('Unexpected live or DLL evidence path')
            arc = item + '/' + path.relative_to(folder).as_posix()
            data = path.read_bytes()
            files[arc] = dict(size=len(data), sha256=hashlib.sha256(data).hexdigest(), source=str(path))
    destination = VARIANT / 'evidence'
    destination.mkdir(exist_ok=True)
    archive = destination / (name + '.zip')
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED, compresslevel=9) as writer:
        for arc, item in files.items():
            info = zipfile.ZipInfo(arc, (2026, 10, 14, 0, 0, 0))
            info.compress_type = zipfile.ZIP_DEFLATED
            info.external_attr = 0o644 << 16
            writer.writestr(info, Path(item['source']).read_bytes(), compresslevel=9)
    manifest = dict(archive=archive.name, size=archive.stat().st_size,
                    sha256=hashlib.sha256(archive.read_bytes()).hexdigest(),
                    files={arc: {k: v for k, v in item.items() if k != 'source'} for arc, item in files.items()},
                    runs=runs, excluded='Workbooks, DLL copies, live production, duplicate tamper payloads',
                    scope='Original matrices/fields/duals and completed strict audits; physical proof only where stated')
    (destination / (name + '_manifest.json')).write_text(json.dumps(manifest, ensure_ascii=False, indent=2) + '\n', encoding='utf-8')
    with zipfile.ZipFile(archive) as reader:
        assert reader.testzip() is None
        assert reader.namelist() == list(files)
        for arc, item in files.items():
            data = reader.read(arc)
            assert len(data) == item['size'] and hashlib.sha256(data).hexdigest() == item['sha256']
    print(json.dumps(dict(archive=str(archive), sha256=manifest['sha256'], bytes=manifest['size'], files=len(files), runs=len(runs))))


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--include', action='append', default=[])
    parser.add_argument('--only', action='append', default=[])
    parser.add_argument('--name', default='native_recovery_primal_evidence_20261014')
    args = parser.parse_args()
    freeze((args.only if args.only else DEFAULTS) + args.include, args.name)
