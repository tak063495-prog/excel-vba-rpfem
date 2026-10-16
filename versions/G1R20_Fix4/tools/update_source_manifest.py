"""Record G1R20 Fix4 bytes and the exact changes from the frozen G1R20 Fix3."""
from pathlib import Path
import hashlib,json
from current_sources import bundle_key
ROOT=Path(__file__).resolve().parents[1]
DLL_SHA='66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b'
BASE=ROOT.parent/'G1R20_Fix3'
CHANGED={'RPX_Guard.bas', 'RPX_Adapt.bas', 'RPX_Seep.bas', 'RPX_SeepUI.bas', 'RPX_Seismic.bas', 'RPX_Control.bas', 'RPX_Mesh.bas', 'RPX_P4.bas', 'RPX_Water.bas', 'RPX_SeepWT.bas', 'RPX_Definition.bas', 'RPX_Model.bas', 'RPX_C0Select.bas', 'RPX_UI.bas'}
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def update():
    modules={}
    for p in sorted((ROOT/'src/vba').glob('*.bas')):
        imported=ROOT/'import'/p.name
        assert p.read_text(encoding='utf8')==imported.read_text(encoding='cp932'),p.name
        modules[p.name]={'utf8_sha256':sha(p),'cp932_sha256':sha(imported)}
    assert len(modules)==63
    changed={name for name in modules if sha(ROOT/'src/vba'/name)!=sha(BASE/'src/vba'/name)}
    assert changed==CHANGED,changed
    workbook_modules={}
    source=ROOT/'src/workbook/ThisWorkbook.code.txt';imported=ROOT/'import/ThisWorkbook.code.txt'
    assert source.read_text(encoding='utf8')==imported.read_text(encoding='cp932')
    workbook_modules['ThisWorkbook']={'utf8_path':source.relative_to(ROOT).as_posix(),
        'cp932_path':imported.relative_to(ROOT).as_posix(),
        'utf8_sha256':sha(source),'cp932_sha256':sha(imported)}
    assets={}
    if (ROOT/'delivery').exists():
        for p in sorted((ROOT/'delivery').glob('*')):
            if p.is_file() and p.suffix.lower() in ['.xlsm','.dll']:
                assets[p.relative_to(ROOT).as_posix()]={'bytes':p.stat().st_size,'sha256':sha(p)}
    data={'schema':'rpfem.g1r20.fix4.sources.v1','version':'G1R20_Fix4','base_version':'G1R20_Fix3',
        'base_manifest_sha256':sha(BASE/'SOURCE_MANIFEST.json'),
        'changed_modules':sorted(changed),
        'base_module_sha256':{name:sha(BASE/'src/vba'/name) for name in sorted(modules)},
        'modules':modules,'standard_module_count':63,'workbook_modules':workbook_modules,'workbook_module_count':1,
        'runtime_dll':'rpfem_native_x64.dll','runtime_dll_sha256':DLL_SHA,
        'assets':assets,'p06_native_complete_claimed':False,
        'scope':'Fix3 review: strict input contracts, integer and binary settings, geometry and legacy text validation, seismic load phase, active seep cancellation and failed producer state; GAP default 2 percent',
        'validation_warning':'Earlier release proofs retain their original source hashes; final Fix4 runs bind all 63 standard modules and the workbook event source; fixed-point wet regression is distinct from a cold root search',
        'source_is_authoritative':'src/vba and src/workbook; UTF8 and CP932 decoded text exactly equal; all release bytes preserved by .gitattributes'}
    data['source_bundle_sha256']=bundle_key(data)
    (ROOT/'SOURCE_MANIFEST.json').write_text(json.dumps(data,indent=2)+'\n',encoding='utf8')
    return {'status':'PASS','modules':len(modules),'assets':len(assets),'manifest_sha256':sha(ROOT/'SOURCE_MANIFEST.json')}
if __name__=='__main__':print(json.dumps(update()))
