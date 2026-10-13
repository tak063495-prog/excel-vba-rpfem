"""Copy independent QA evidence and record this exact DLL's PE dependencies."""
from pathlib import Path
import argparse
import hashlib
import json
import shutil
import subprocess


def sha(path):
    return hashlib.sha256(path.read_bytes()).hexdigest()


def run(qa_path, dumpbin):
    root = Path(__file__).resolve().parent
    dll = root / 'bin/rpfem_accuracy_x64.dll'
    build = json.loads((root / 'bin/build_manifest.json').read_text(encoding='utf-8-sig'))
    original_sha = sha(dll)
    if build['dll_sha256'] != original_sha:
        raise ValueError('DLL does not match its build manifest')
    qa = json.loads(qa_path.read_text(encoding='utf8'))
    if not qa['sources_unchanged'] or qa['protected_sha256'] != qa['protected_sha256_after']:
        raise ValueError('Independent QA source/binary freeze failed')
    if qa['protected_sha256'].get(str(dll)) != original_sha:
        raise ValueError('Independent QA belongs to a different DLL')
    outputs = {}
    for option in ('DEPENDENTS', 'EXPORTS'):
        outputs[option] = subprocess.run([str(dumpbin), '/'+option, str(dll)],
            check=True, capture_output=True, text=True).stdout
    if 'VCRUNTIME' in outputs['DEPENDENTS'].upper() or 'API-MS-WIN-CRT' in outputs['DEPENDENTS'].upper():
        raise ValueError('Static CRT DLL still has redistributable runtime dependencies')
    if sha(dll) != original_sha:
        raise ValueError('DLL changed during PE validation')
    destination = root / 'validation'
    destination.mkdir(exist_ok=True)
    copied = destination / 'independent_ffi_q8_qa.json'
    shutil.copyfile(qa_path, copied)
    if sha(copied) != sha(qa_path):
        raise ValueError('QA evidence copy differs')
    for option, value in outputs.items():
        (destination / ('pe_'+option.lower()+'.txt')).write_text(value, encoding='utf8')
    (destination / 'validation_manifest.json').write_text(json.dumps({
        'dll_sha256': original_sha, 'clarabel': '0.11.1',
        'target': build['target'], 'independent_qa_sha256': sha(copied),
        'pe_inspection_source': str(dumpbin),
        'static_crt_imports_verified': True,
        'scope': 'Small actual original q8 programs and FFI guards; P06 remains outstanding',
        'physical_audits_independent': True, 'solver_status_is_proof': False,
        'modified_qp_constraint_dual_is_original_guide': False,
    }, indent=2)+'\n', encoding='utf8')
    print('RECORDED_NATIVE_VALIDATION', original_sha)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--qa', type=Path, required=True)
    parser.add_argument('--dumpbin', type=Path, required=True)
    arguments = parser.parse_args()
    run(arguments.qa, arguments.dumpbin)
