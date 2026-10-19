"""Refresh SOURCE_MANIFEST.json for Fix9 from src/vba (ASCII sources: import copy is byte-identical)."""
from pathlib import Path
import hashlib,json,shutil
ROOT=Path(__file__).resolve().parents[1];BASE=ROOT.parent/'G1R20_Fix8'
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
m.update(schema='rpfem.g1r20.fix9.sources.v1',version='G1R20_Fix9_candidate',base_version='G1R20_Fix8',
    base_manifest_sha256=sha(BASE/'SOURCE_MANIFEST.json'),changed_modules=changed,
    base_module_sha256={k:v['utf8_sha256'] for k,v in base['modules'].items()},
    scope='G1R20 Fix9 candidate: Fix8 thin-layer chains aligned with the Python prototype (separate band mid-line step with own root, lower indicator from an upper solve on the lower mesh, fresh upper solve before flips, lower retry with smaller theta after one root failure, chain-only HSD budget scaling by cone count and DLL-direct upper time scaling); all non-thin-layer paths identical to Fix8',
    validation_warning='Candidate; Winmax2 regression P10A/D05A/P02A/P05A/P04A/D07A and forced thin-layer smoke P02A_THIN only; P06A/P06N not run in VBA; not published')
payload={k:m[k] for k in ('modules','workbook_modules')}
m['source_bundle_sha256']=hashlib.sha256(json.dumps(payload,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
(ROOT/'SOURCE_MANIFEST.json').write_text(json.dumps(m,indent=2,ensure_ascii=False)+'\n',encoding='utf8')
print('changed',changed)
