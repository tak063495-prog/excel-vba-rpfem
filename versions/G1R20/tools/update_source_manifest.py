"""Record exact G1R20 release bytes. No workbook/solver is opened."""
from pathlib import Path
import hashlib,json
ROOT=Path(__file__).resolve().parents[1]
DLL_SHA='66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b'
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def update():
    modules={}
    for p in sorted((ROOT/'src/vba').glob('*.bas')):
        imported=ROOT/'import'/p.name
        assert p.read_text(encoding='utf8')==imported.read_text(encoding='cp932'),p.name
        modules[p.name]={'utf8_sha256':sha(p),'cp932_sha256':sha(imported)}
    assert len(modules)==63
    assets={}
    if (ROOT/'delivery').exists():
        for p in sorted((ROOT/'delivery').glob('*')):
            if p.is_file() and p.suffix.lower() in ['.xlsm','.dll']:
                assets[p.relative_to(ROOT).as_posix()]={'bytes':p.stat().st_size,'sha256':sha(p)}
    data={'schema':'rpfem.g1r20.sources.v1','version':'G1R20','development_candidate':True,
        'modules':modules,'standard_module_count':63,'runtime_dll':'rpfem_native_x64.dll','runtime_dll_sha256':DLL_SHA,
        'assets':assets,'p06_native_complete_claimed':False,
        'scope':'Combined10 plus FlowMesh, optional balanced-red4/flow candidates, terminal schedule and bounded same-Fs upper recovery',
        'validation_warning':'Previous four/seven/ten-module proofs retain their own source hashes; geometry replay is distinct from automatic cold generation',
        'source_is_authoritative':'src/vba; UTF8 and CP932 decoded text exactly equal; all release bytes preserved by .gitattributes'}
    (ROOT/'SOURCE_MANIFEST.json').write_text(json.dumps(data,indent=2)+'\n',encoding='utf8')
    return {'status':'PASS','modules':len(modules),'assets':len(assets),'manifest_sha256':sha(ROOT/'SOURCE_MANIFEST.json')}
if __name__=='__main__':print(json.dumps(update()))
