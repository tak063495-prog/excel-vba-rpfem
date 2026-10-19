"""Refresh SOURCE_MANIFEST.json for Fix9 rc4b from src/vba (ASCII sources: import copy is byte-identical)."""
from pathlib import Path
import hashlib,json,shutil
ROOT=Path(__file__).resolve().parents[1];BASE=ROOT.parent/'G1R20_Fix9rc2'
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
base=json.loads((BASE/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
m=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
changed=[]
for name in sorted(base['modules']):
    src=ROOT/'src/vba'/name
    if sha(src)==base['modules'][name]['utf8_sha256']:
        m['modules'][name]=dict(base['modules'][name]);assert sha(ROOT/'import'/name)==base['modules'][name]['cp932_sha256'];continue
    src.read_bytes().decode('ascii')  # changed modules must stay ASCII (CP932 copy identical)
    shutil.copyfile(src,ROOT/'import'/name)
    m['modules'][name]={'utf8_sha256':sha(src),'cp932_sha256':sha(ROOT/'import'/name)};changed.append(name)
m.update(schema='rpfem.g1r20.fix9rc4b.sources.v1',version='G1R20_Fix9rc4b_candidate',base_version='G1R20_Fix9rc2',
    base_manifest_sha256=sha(BASE/'SOURCE_MANIFEST.json'),changed_modules=changed,
    base_module_sha256={k:v['utf8_sha256'] for k,v in base['modules'].items()},
    scope='G1R20 Fix9 rc4b candidate (rc4 with ADAPT_THIN_EARLY_UPPER_THETA default 0.95) on top of rc3b (Codex-written, reviewed): after an early start is taken, upper chain uses ADAPT_THIN_EARLY_UPPER_THETA=0.95 and ADAPT_THIN_EARLY_UPPER_CAP=7000 (budget check uses them too); otherwise identical to rc3b. rc3b: conditional early start ADAPT_THIN_EARLY_START=1 (Fix9ThinEarlyBudget: min upper rounds 10, min lower rounds 3, lower end <= upper rounds, theta >= 0.8/0.9, cap/coarse >= 2, remaining optional deadline >= 3600 s if active; 2 = always as rc2, 0 = off); ADAPT_THIN_AUX_ALMOST=1; lower cap 6000',
    validation_warning='Candidate; Winmax2 regression P10A/D05A/P02A/P05A/P04A/D07A and forced thin-layer P02A_THIN only; P06A/P06N not run in VBA; not published')
payload={k:m[k] for k in ('modules','workbook_modules')}
m['source_bundle_sha256']=hashlib.sha256(json.dumps(payload,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
(ROOT/'SOURCE_MANIFEST.json').write_text(json.dumps(m,indent=2,ensure_ascii=False)+'\n',encoding='utf8')
print('changed',changed)
