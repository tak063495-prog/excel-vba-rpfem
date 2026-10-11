from pathlib import Path
import hashlib,json
ROOT=Path(__file__).resolve().parents[1]; base=ROOT.parent.parent
files=json.loads((ROOT/'results/protected_previous_delivery.json').read_text(encoding='utf-8'))
broken=[p for p,h in files.items() if hashlib.sha256((base/p).read_bytes()).hexdigest()!=h]
assert not broken,broken
print('PASS previous DLL delivery and ZIP preserved:',len(files),'files')
