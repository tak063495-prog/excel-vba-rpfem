"""CI source/import/manifest and optional delivery-byte integrity checks."""
from pathlib import Path
import argparse,hashlib,json,re,zipfile
from current_sources import bundle_key
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
    assert manifest['version']=='G1R20_Fix3' and manifest['base_version']=='G1R20_Fix2'
    workbook_modules=manifest['workbook_modules']
    assert set(workbook_modules)=={'ThisWorkbook'} and manifest['workbook_module_count']==1
    for name,row in workbook_modules.items():
        assert row['utf8_path']=='src/workbook/ThisWorkbook.code.txt'
        assert row['cp932_path']=='import/ThisWorkbook.code.txt'
        a=(ROOT/row['utf8_path']).read_bytes();b=(ROOT/row['cp932_path']).read_bytes()
        assert sha(a)==row['utf8_sha256'] and sha(b)==row['cp932_sha256'],name
        assert a.decode('utf8')==b.decode('cp932'),name
    assert bundle_key(manifest)==manifest['source_bundle_sha256']
    base=ROOT.parent/manifest['base_version']
    assert sha((base/'SOURCE_MANIFEST.json').read_bytes())==manifest['base_manifest_sha256']
    changed=[]
    for name,p in sources.items():
        baseline=sha((base/'src/vba'/name).read_bytes())
        assert baseline==manifest['base_module_sha256'][name],name
        if sha(p.read_bytes())!=baseline:changed.append(name)
    assert sorted(changed)==manifest['changed_modules']==['RPX_Control.bas','RPX_Fs.bas','RPX_Guard.bas','RPX_Model.bas','RPX_Seep.bas','RPX_SeepUI.bas','RPX_Water.bas']
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
        validation_path=ROOT/'delivery/WORKBOOK_VALIDATION.json'
        audit_path=ROOT/'SOURCE_EXTRACTION_AUDIT.json'
        if require_delivery:
            assert validation_path.exists() and audit_path.exists(),'Workbook validation and independent extraction audit required'
        if validation_path.exists():
            validation=json.loads(validation_path.read_text(encoding='utf8'))
            assert validation['state']=='PASS' and validation['owned_excel_exited']
            assert validation['workbook_sha256']==sha(book.read_bytes())
            assert validation['source_bundle_sha256']==manifest['source_bundle_sha256']
            assert validation['standard_modules']==63 and validation['workbook_module_count']==1
            assert set(validation['modules'])==set(manifest['modules'])
            for name,row in manifest['modules'].items():
                assert validation['modules'][name]['source_sha256']==row['utf8_sha256'],name
            for name,row in workbook_modules.items():
                assert validation['workbook_modules'][name]['source_sha256']==row['utf8_sha256'],name
        if audit_path.exists():
            audit=json.loads(audit_path.read_text(encoding='utf8'))
            assert audit['status']=='PASS' and audit['workbook_sha256']==sha(book.read_bytes())
            assert audit['source_bundle_sha256']==manifest['source_bundle_sha256']
            assert audit['module_count']==63 and audit['workbook_module_count']==1
            assert set(audit['modules'])==set(manifest['modules']) and set(audit['workbook_modules'])==set(workbook_modules)
            for name,row in manifest['modules'].items():
                assert audit['modules'][name]['source_sha256']==row['utf8_sha256'] and audit['modules'][name]['pass'],name
            for name,row in workbook_modules.items():
                assert audit['workbook_modules'][name]['source_sha256']==row['utf8_sha256'] and audit['workbook_modules'][name]['pass'],name
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
            for name,row in workbook_modules.items():
                for path_key,hash_key in [('utf8_path','utf8_sha256'),('cp932_path','cp932_sha256')]:
                    matches=[n for n in names if n.endswith('/'+row[path_key]) or n==row[path_key]]
                    assert len(matches)==1 and sha(z.read(matches[0]))==row[hash_key],(package.name,name,path_key)
            for audit_name in ['SOURCE_EXTRACTION_AUDIT.json','delivery/WORKBOOK_VALIDATION.json']:
                local=ROOT/audit_name
                if local.exists():
                    matches=[n for n in names if n.endswith('/'+audit_name) or n==audit_name]
                    assert len(matches)==1 and sha(z.read(matches[0]))==sha(local.read_bytes()),(package.name,audit_name)
            for name,row in assets.items():
                matches=[n for n in names if n.endswith('/'+Path(name).name) or n==Path(name).name]
                assert len(matches)==1 and sha(z.read(matches[0]))==row['sha256'],(package.name,name)
    return {'status':'PASS','modules':len(sources),'workbook_modules':len(workbook_modules),'assets':len(assets),'workbooks':len(books),'runtime_dlls':len(dlls),'release_zips':len(release_zips),'source_bundle_sha256':manifest['source_bundle_sha256'],'manifest_sha256':sha((ROOT/'SOURCE_MANIFEST.json').read_bytes())}
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--require-delivery',action='store_true');print(json.dumps(check(p.parse_args().require_delivery)))
