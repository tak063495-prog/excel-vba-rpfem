"""CI source/import/manifest and optional delivery-byte integrity checks."""
from pathlib import Path
import argparse,hashlib,json,re,zipfile
ROOT=Path(__file__).resolve().parents[1]
def sha(b):return hashlib.sha256(b).hexdigest()
def check(require_delivery=False):
    manifest=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    sources={p.name:p for p in (ROOT/'src/vba').glob('*.bas')};imports={p.name:p for p in (ROOT/'import').glob('*.bas')}
    assert len(sources)==manifest['standard_module_count']==63
    assert set(sources)==set(imports)==set(manifest['modules'])
    for name,p in sources.items():
        a=p.read_bytes();b=imports[name].read_bytes();row=manifest['modules'][name]
        assert sha(a)==row['utf8_sha256'] and sha(b)==row['cp932_sha256'],name
        assert a.decode('utf8')==b.decode('cp932'),name
        text=a.decode('utf8');assert re.search(r'^Attribute VB_Name = "'+re.escape(p.stem)+r'"\s*$',text,re.M),name
        assert 'rpfem_accuracy_x64.dll' not in text.lower(),name
    base=ROOT.parent/manifest['base_version']
    assert sha((base/'SOURCE_MANIFEST.json').read_bytes())==manifest['base_manifest_sha256']
    changed=[]
    for name,p in sources.items():
        baseline=sha((base/'src/vba'/name).read_bytes())
        assert baseline==manifest['base_module_sha256'][name],name
        if sha(p.read_bytes())!=baseline:changed.append(name)
    assert sorted(changed)==manifest['changed_modules']==['RPX_Seep.bas','RPX_SeepUI.bas','RPX_UI.bas','RPX_Water.bas']
    assets=manifest.get('assets',{})
    for name,row in assets.items():
        p=ROOT/name;raw=p.read_bytes();assert len(raw)==row['bytes'] and sha(raw)==row['sha256'],name
    dlls=list((ROOT/'delivery').glob('*.dll')) if (ROOT/'delivery').exists() else []
    books=list((ROOT/'delivery').glob('*.xlsm')) if (ROOT/'delivery').exists() else []
    if require_delivery:
        assert len(dlls)==1 and len(books)==1,'Require one runtime and one workbook'
    if dlls:
        assert len(dlls)==1 and dlls[0].name==manifest['runtime_dll']
        assert sha(dlls[0].read_bytes())==manifest['runtime_dll_sha256']
    for book in books:
        assert book.relative_to(ROOT).as_posix() in assets,'Workbook not bound by source manifest'
        with zipfile.ZipFile(book) as z:assert 'xl/vbaProject.bin' in z.namelist()
    release_zips=[p for p in (ROOT/'delivery').glob('*.zip')] if (ROOT/'delivery').exists() else []
    for package in release_zips:
        with zipfile.ZipFile(package) as z:
            names=z.namelist();inside=[n for n in names if n.lower().endswith('.dll')]
            assert len(inside)==1 and Path(inside[0]).name==manifest['runtime_dll']
            assert sha(z.read(inside[0]))==manifest['runtime_dll_sha256']
            for name,row in manifest['modules'].items():
                for sub,key in [('src/vba','utf8_sha256'),('import','cp932_sha256')]:
                    matches=[n for n in names if n.endswith(sub+'/'+name)]
                    assert len(matches)==1 and sha(z.read(matches[0]))==row[key],(package.name,sub,name)
            for name,row in assets.items():
                matches=[n for n in names if n.endswith('/'+Path(name).name) or n==Path(name).name]
                assert len(matches)==1 and sha(z.read(matches[0]))==row['sha256'],(package.name,name)
    return {'status':'PASS','modules':len(sources),'assets':len(assets),'workbooks':len(books),'runtime_dlls':len(dlls),'release_zips':len(release_zips),'manifest_sha256':sha((ROOT/'SOURCE_MANIFEST.json').read_bytes())}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--require-delivery',action='store_true');print(json.dumps(check(p.parse_args().require_delivery)))
