"""Freeze completed native component proofs without mixing source scopes."""
from pathlib import Path
import argparse
import hashlib
import json
import zipfile

REPO = Path(__file__).resolve().parents[2]
SCOPES = {
    'seven_module_D08_cold': 'validation/G1R19_WaterCommon_excel_20261016/results/owned_excel_1792152100656896200',
    'water_transaction_micro_solver0': 'validation/G1R19_WaterCommon_excel_20261016/results/owned_micro_1792151829932868800',
    'state_reset_positive_micro_solver0': 'validation/G1R19_StateReset_excel_20261016/results/positive_/1792153293517482300',
    'state_reset_negative_control_solver0': 'validation/G1R19_StateReset_excel_20261016/results/negative_/1792153314267480200',
    'ten_module_D02_cold': 'validation/G1R19_Combined_excel_20261016/results/owned_excel_1792153483764082600',
    'flow_geometry_micro_solver0': 'validation/G1R19_FlowMesh_excel_20261016/results/owned_geometry_1792154199526082100',
}


def digest(data):
    return hashlib.sha256(data).hexdigest()


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument('--output', type=Path, required=True)
    args = ap.parse_args()
    if args.output.exists():
        raise FileExistsError('Existing evidence is immutable')
    paths = set()
    scopes = []
    for role, name in SCOPES.items():
        folder = REPO/name
        summary = json.loads((folder/'summary.json').read_text(encoding='utf8'))
        if summary.get('status') != 'PASS' or summary.get('owned_excel_exited') is not True:
            raise ValueError('Component did not finish successfully: '+role)
        scopes.append({'role': role, 'folder': name, 'status': 'PASS',
                       'summary_sha256': digest((folder/'summary.json').read_bytes())})
        paths.update(p for p in folder.rglob('*') if p.is_file() and '__pycache__' not in p.parts)
    for name in ['validation/G1R19_D08_native_audit_20261016',
                 'validation/G1R19_D08_ffi_20261016',
                 'validation/G1R19_WaterCommon_excel_20261016',
                 'validation/G1R19_StateReset_excel_20261016',
                 'validation/G1R19_FlowMesh_excel_20261016',
                 'validation/G1R19_Combined_excel_20261016']:
        folder = REPO/name
        paths.update(p for p in folder.glob('*') if p.is_file() and p.suffix in ('.py', '.md', '.bas'))
    paths.add(REPO/'validation/G1R19_D02_route_audit_20261016/COLD_COMBINED_D02_INDEPENDENT_AUDIT.json')
    paths.add(Path(__file__).resolve())
    records = {p.relative_to(REPO).as_posix(): digest(p.read_bytes()) for p in sorted(paths)}
    manifest = {'status': 'COMPLETED_COMPONENT_PROOFS', 'full_P06_native_claimed': False,
                'full_26_native_claimed': False, 'source_scopes_are_separate': True,
                'scopes': scopes, 'files': records,
                'note': 'Original paths and bytes retained; micro tests are not numerical bound proofs.'}
    args.output.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(args.output, 'x', zipfile.ZIP_DEFLATED, compresslevel=6) as archive:
        for name, expected in records.items():
            data = (REPO/name).read_bytes()
            if digest(data) != expected:
                raise ValueError('Evidence changed during packing: '+name)
            archive.writestr(name, data)
        archive.writestr('NATIVE_COMPONENT_MANIFEST.json', json.dumps(manifest, ensure_ascii=False, indent=2)+'\n')
    with zipfile.ZipFile(args.output) as archive:
        if archive.testzip() is not None or any(digest(archive.read(n)) != h for n, h in records.items()):
            raise ValueError('Archive CRC or member SHA failed')
    record = {'status': 'PASS', 'scope': manifest['status'], 'archive': args.output.name,
              'file_count': len(records), 'bytes': args.output.stat().st_size,
              'sha256': digest(args.output.read_bytes()), 'all_member_sha256_and_crc': 'PASS', 'scopes': scopes}
    args.output.with_suffix('.json').write_text(json.dumps(record, ensure_ascii=False, indent=2)+'\n', encoding='utf8')
    print(json.dumps({k: v for k, v in record.items() if k != 'scopes'}))


if __name__ == '__main__':
    main()
