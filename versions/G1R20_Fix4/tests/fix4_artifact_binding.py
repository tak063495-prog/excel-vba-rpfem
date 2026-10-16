"""Bind test source snapshots to the independently built/saved release workbook."""
from pathlib import Path
import hashlib, json

def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def bind(root):
    root=Path(root);validation=root/'delivery/WORKBOOK_VALIDATION.json'
    report=json.loads(validation.read_text(encoding='utf8'))
    book=root/'delivery/RPFEM_20261017_G1R20_Fix4.xlsm';dll=root/'delivery/rpfem_native_x64.dll'
    paths=sorted((root/'src/vba').glob('*.bas'));doc=root/'src/workbook/ThisWorkbook.code.txt'
    assert report['state']=='PASS' and report['owned_excel_exited']
    assert len(paths)==63 and report['standard_modules']==63 and report['workbook_module_count']==1
    assert report['workbook_sha256']==sha(book)
    assert report['dll_sha256']==sha(dll)=='66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b'
    for path in paths:assert report['modules'][path.name]['source_sha256']==sha(path),path.name
    assert report['workbook_modules']['ThisWorkbook']['source_sha256']==sha(doc)
    return dict(workbook_sha256=sha(book),dll_sha256=sha(dll),validation_sha256=sha(validation),source_bundle_sha256=report['source_bundle_sha256'],standard_modules=63,document_modules=1,all_64_source_sha256_match=True)
