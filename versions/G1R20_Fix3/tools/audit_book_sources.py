"""Offline check of 63 saved standard modules and ThisWorkbook against sources."""
from pathlib import Path
import argparse,hashlib,importlib.util,json
from oletools.olevba import VBA_Parser
from current_sources import bundle_key

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('combined_builder',ROOT/'tools/build_candidate_book.py')
builder=importlib.util.module_from_spec(spec);spec.loader.exec_module(builder)
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def key(text):
    lines=[s for s in text.replace('\r\n','\n').splitlines() if not s.startswith('Attribute VB_')]
    return builder.code_key('\r\n'.join(lines))

def audit(output):
    output=Path(output).resolve();book=output/'delivery/RPFEM_20261017_G1R20_Fix3.xlsm'
    before=sha(book);parser=VBA_Parser(str(book));sources={p.name:p for p in (output/'src/vba').glob('*.bas')};rows={};workbook_rows={}
    manifest=json.loads((output/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    assert bundle_key(manifest)==manifest['source_bundle_sha256']
    assert set(sources)==set(manifest['modules'])
    for name,source in sources.items():assert sha(source)==manifest['modules'][name]['utf8_sha256'],name
    try:
        for _,stream,name,text in parser.extract_macros():
            if isinstance(text,bytes):text=text.decode('cp932')
            if name=='ThisWorkbook.cls':
                row=manifest['workbook_modules']['ThisWorkbook'];source=output/row['utf8_path']
                assert sha(source)==row['utf8_sha256']
                want=key(source.read_text(encoding='utf8'));actual=key(text)
                workbook_rows['ThisWorkbook']={'stream':stream,'extracted_name':name,'source_sha256':sha(source),
                    'source_code_key_sha256':want,'extracted_code_key_sha256':actual,'pass':want==actual}
                continue
            if name not in sources:continue
            source=sources[name];want=key(source.read_text(encoding='utf8'));actual=key(text)
            rows[name]={'stream':stream,'source_sha256':sha(source),'source_code_key_sha256':want,
                        'extracted_code_key_sha256':actual,'pass':want==actual}
    finally:parser.close()
    assert set(rows)==set(sources) and len(rows)==63
    assert all(row['pass'] for row in rows.values()),[n for n,r in rows.items() if not r['pass']]
    assert set(workbook_rows)=={'ThisWorkbook'} and all(row['pass'] for row in workbook_rows.values()),workbook_rows
    assert before==sha(book)
    data={'status':'PASS','scope':'Saved XLSM source extraction, no Excel or solver',
          'workbook_sha256':before,'module_count':63,'workbook_module_count':1,'source_bundle_sha256':manifest['source_bundle_sha256'],'verifier_sha256':sha(__file__),
          'builder_sha256':sha(ROOT/'tools/build_candidate_book.py'),
          'normalization':'Identifier case and leading-label indentation only; literal strings/comments/numbers/Call preserved',
          'modules':rows,'workbook_modules':workbook_rows}
    path=output/'SOURCE_EXTRACTION_AUDIT.json';path.write_text(json.dumps(data,indent=2)+'\n',encoding='utf8')
    print(json.dumps({'status':'PASS','module_count':63,'audit':str(path),'workbook_sha256':before}))

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('output',type=Path);audit(parser.parse_args().output)
