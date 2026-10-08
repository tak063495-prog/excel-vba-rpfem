from pathlib import Path
import hashlib
import json
import subprocess
import sys

R = Path(__file__).resolve().parent
baseline = Path('C:/Users/link_/Downloads/RPFEM_Codex_P2S_0945_Task')
proc = subprocess.run([sys.executable, '-X', 'utf8', 'tools/verify_baseline.py'],
                      cwd=baseline, text=True, encoding='utf-8', capture_output=True,
                      check=True)
result = json.loads(proc.stdout)
assert result['baseline_intact'] and result['checked_files'] == 32
(R / 'results/baseline_after.json').write_text(
    json.dumps(result, indent=2), encoding='utf-8')
identity = json.loads((R / 'source_identity.json').read_text(encoding='utf-8'))
files = [identity,
         {'path': 'C:/Users/link_/Desktop/RPFEM_G1R9_Speedup_20261006/RPFEM_20261006_G1R9_P06N_DD_Fast_Verified.xlsm',
          'sha256': '3f6a405ae05bd65dfa35bd3bf6413b6d43f51dc8298eabb16897d1a9e91ed1f0'}]
for item in files:
    assert hashlib.sha256(Path(item['path']).read_bytes()).hexdigest() == item['sha256'], item
result = {'baseline_reference_32_files': 'PASS', 'original_workbooks': files,
          'original_workbooks_intact': True}
(R / 'results/protected_originals.json').write_text(
    json.dumps(result, indent=2), encoding='utf-8')
print('Protected baseline/reference and G1R9/G1R10 originals PASS')
