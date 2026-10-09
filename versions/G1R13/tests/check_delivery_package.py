"""Extract and independently replay the actual delivery ZIP, without original evidence paths."""
from pathlib import Path
import hashlib, json, subprocess, sys, zipfile
ROOT = Path(__file__).resolve().parents[1]
package = ROOT / 'delivery/RPFEM_G1R13_D03N_D07A_Checked.zip'
destination = ROOT / 'delivery/unpacked_check'
assert not destination.exists(), destination
destination.mkdir()
with zipfile.ZipFile(package) as archive:
    assert archive.testzip() is None
    for item in archive.infolist():
        target = (destination / item.filename).resolve()
        assert target.is_relative_to(destination.resolve()), item.filename
        assert not item.filename.startswith(('/', '\\')), item.filename
    archive.extractall(destination)
manifest = json.loads((destination / 'MANIFEST.json').read_text(encoding='utf-8'))
for name, digest in manifest.items():
    assert hashlib.sha256((destination / name).read_bytes()).hexdigest() == digest, name
replay = subprocess.run([sys.executable, '-X', 'utf8', str(destination / 'tests/verify_portable.py')], cwd=destination, text=True, encoding='utf-8', capture_output=True)
assert replay.returncode == 0, (replay.stdout, replay.stderr)
original = json.loads((ROOT / 'results/portable_audit.json').read_text(encoding='utf-8'))
extracted = json.loads((destination / 'results/portable_audit.json').read_text(encoding='utf-8'))
assert original == extracted
assert len(extracted['results']) == 7
record = {'pass': True, 'zip_sha256': hashlib.sha256(package.read_bytes()).hexdigest(), 'manifest_files_verified': len(manifest), 'extracted_replay': json.loads(replay.stdout), 'replay_result_exactly_matches': True, 'extracted_directory': str(destination)}
(ROOT / 'delivery/EXTRACTION_VERIFIED.json').write_text(json.dumps(record, ensure_ascii=False, indent=2), encoding='utf-8')
print(json.dumps(record, ensure_ascii=False))
