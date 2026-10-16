"""Package checked source, workbook and exactly one DLL without changing them."""
from pathlib import Path
import argparse, hashlib, json, zipfile
import check_release_sources

ROOT = Path(__file__).resolve().parents[1]


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--output', type=Path, required=True)
    args = parser.parse_args()
    if args.output.exists():
        raise FileExistsError('Published package bytes must not be replaced')
    check_release_sources.check(require_delivery=True)
    paths = {ROOT/name for name in ['README_JA.md','TEST_RESULTS.md','PYTHON_VALIDATION.md',
                                    'SOURCE_MANIFEST.json','SOURCE_EXTRACTION_AUDIT.json']}
    for directory in ('src/vba','import','tools','tests'):
        paths.update(p for p in (ROOT/directory).rglob('*') if p.is_file() and
                     not {'results','__pycache__'}.intersection(p.parts) and p.suffix in ('.py','.bas','.txt'))
    paths.update(p for p in (ROOT/'delivery').glob('*') if p.suffix in ('.xlsm','.dll') or
                 p.name == 'WORKBOOK_VALIDATION.json')
    files = {p.relative_to(ROOT).as_posix():p.read_bytes() for p in sorted(paths)}
    dlls = [name for name in files if name.lower().endswith('.dll')]
    if dlls != ['delivery/rpfem_native_x64.dll']:
        raise ValueError('Runtime ZIP must contain exactly one DLL')
    sums = ''.join(sha(raw)+'  '+name+'\n' for name, raw in sorted(files.items()))
    (ROOT/'SHA256SUMS.txt').write_bytes(sums.encode('utf8'))
    files['SHA256SUMS.txt'] = sums.encode('utf8')
    prefix = 'RPFEM_20261016_G1R20/'
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(args.output, 'x', zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for name, raw in sorted(files.items()):
            archive.writestr(prefix+name, raw)
    with zipfile.ZipFile(args.output) as archive:
        if archive.testzip() is not None or any(sha(archive.read(prefix+name)) != sha(raw) for name,raw in files.items()):
            raise ValueError('Packaged bytes or CRC differ')
    record = dict(status='PASS', version='G1R20', archive=args.output.name,
                  bytes=args.output.stat().st_size, sha256=sha(args.output.read_bytes()),
                  member_count=len(files), runtime_dll_count=1, source_module_count=63,
                  workbook_sha256=sha(files['delivery/RPFEM_20261016_G1R20.xlsm']),
                  member_sha256_and_crc='PASS', p06_native_cold_completion_claimed=False)
    args.output.with_suffix('.json').write_text(json.dumps(record, indent=2)+'\n', encoding='utf8')
    print(json.dumps(record))


if __name__ == '__main__':
    main()
