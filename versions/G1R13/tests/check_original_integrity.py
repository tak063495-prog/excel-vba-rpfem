from pathlib import Path
import hashlib, json, subprocess
ROOT = Path(__file__).resolve().parents[1]
expected = {
    'C:/Users/link_/Downloads/excel-vba-rpfem_G1R11_20261008/versions/G1R12/RPFEM_20261009_G1R12_P06N_IssueFix.xlsm': 'fe02ff61687cff8e0f0fc775867c8a244a83f758a3c5b9a9f80b6a2ab0389520',
    'C:/Users/link_/Desktop/RPFEM_G1R12_Native25_20261009/queued_continuation/cases/D03N/G1R12_D03N_Full1536_Q8.xlsm': 'f5f1555306675b7b72c6cc242029abcb1cdc539e0492bfe88ec4d6232fd75605',
}
checks = []
for name, digest in expected.items():
    actual = hashlib.sha256(Path(name).read_bytes()).hexdigest()
    assert actual == digest, name
    checks.append({'file': name, 'sha256': actual, 'unchanged': True})
git_status = subprocess.check_output(['git', 'status', '--short'], cwd='C:/Users/link_/Downloads/excel-vba-rpfem_G1R11_20261008', text=True)
assert not git_status.strip(), git_status
record = {'pass': True, 'original_workbooks': checks, 'repository_worktree_clean': True}
(ROOT / 'results/original_integrity.json').write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(record, ensure_ascii=False))
