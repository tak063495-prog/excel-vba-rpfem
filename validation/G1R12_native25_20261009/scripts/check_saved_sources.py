"""Read saved VBA sources only; never edit XLSM or rebuild vbaProject.bin."""
from pathlib import Path
import json,re,hashlib
from oletools.olevba import VBA_Parser
from prepare_inputs import ROOT,SOURCE,REPO,SOURCE_SHA
from case_locations import case_folder,evidence_root

def normalize(text):
    rows=[]
    for line in text.replace('\r','').splitlines():
        if line.startswith('Attribute VB_') or not line.strip():continue
        parts=re.split(r'("(?:[^" ]| |"")*"|\x27.*)',line.rstrip())
        rows.append(''.join(p if i%2 else p.casefold() for i,p in enumerate(parts)))
    return '\n'.join(rows)

def modules(path):
    parser=VBA_Parser(str(path))
    try:return {name:normalize(code) for _,_,name,code in parser.extract_macros()}
    finally:parser.close()

def check(case,base=None):
    folder=case_folder(case);record=json.loads((folder/'status.json').read_text(encoding='utf-8'))
    assert record['status'] not in ['RUNNING','PREPARING','NOT_RUN'] and record.get('workbook')
    base=modules(SOURCE) if base is None else base
    actual=modules(evidence_root(case)/record['workbook'])
    assert set(actual)==set(base)|{'BatchEvidence.bas'},(case,'unexpected module set')
    for name,expected in base.items():assert actual[name]==expected,(case,name,'source changed')
    source_files=list((REPO/'versions/G1R12/src').glob('*.bas'));assert len(source_files)==46
    for p in source_files:assert actual[p.name]==normalize(p.read_text(encoding='utf-8')),(case,p.name,'production source discrepancy')
    assert actual['BatchEvidence.bas']==normalize((ROOT/'tests/BatchEvidence.bas').read_text(encoding='utf-8'))
    assert hashlib.sha256(SOURCE.read_bytes()).hexdigest()==SOURCE_SHA
    result={'case':case,'pass':True,'original_modules_unchanged':len(base),'production_standard_modules_matched':46,
            'document_modules_unchanged':len([n for n in base if n.endswith('.cls')]),'only_added_module':'BatchEvidence',
            'scope':'read-only VBA source comparison; template compile and native execution are reported separately'}
    (folder/'saved_sources_check.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    return result

if __name__=='__main__':
    base=modules(SOURCE);rows=[]
    for p in sorted((ROOT/'inputs').glob('*.json')):
        status=case_folder(p.stem)/'status.json'
        if not status.exists():continue
        r=json.loads(status.read_text(encoding='utf-8'))
        if r['status'] not in ['RUNNING','PREPARING','NOT_RUN'] and r.get('workbook'):rows.append(check(p.stem,base))
    print(json.dumps(rows,ensure_ascii=False))
