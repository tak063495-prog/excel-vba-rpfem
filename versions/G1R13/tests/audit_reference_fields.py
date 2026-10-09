"""Re-audit prior native results; this is not a new G1R13 Excel solve."""
from pathlib import Path
import sys,json
ROOT=Path(__file__).resolve().parents[1]
OLD=Path('C:/Users/link_/Desktop/RPFEM_G1R12_Native25_20261009')
sys.path[:0]=[str(ROOT/'validation_runtime/python'),str(ROOT/'validation_runtime/vendor'),str(OLD)]
from snapshot_reader import read_state
from audit_precision import audit_precision
from case_locations import case_folder
rows=[]
for case in ['D01A','D02A','P06A','P06N','P02A','P02N']:
    rec=json.loads((ROOT/'inputs'/f'{case}.json').read_text(encoding='utf-8'))
    for side in ['lower','upper']:
        m,x,meta=read_state(case_folder(case)/f'witness_{side}.bin',rec['source_case']['policy'],'reference_total');m.provenance['case']=case
        a=audit_precision(m,x,meta['fs'],side)
        if side=='lower':
            ok=max((v for k,v in a['raw_errors'].items() if k!='load_one'),default=0.)<=1e-7 and a['load_factor']>=1.-1e-8
            contract='lambda_max endpoint: lambda >= 1 - 1e-8; physical raw errors <= 1e-7; not a fixed-unit-load witness'
        else:ok=a['accepted'];contract='upper collapse / physical work audit'
        rows.append({'case':case,'side':side,'metadata':meta,'audit':a,'bound_contract':contract,'bound_contract_pass':ok})
        print(case,side,ok,a['raw_errors'],flush=True)
        assert ok,(case,side,a)
(ROOT/'results/reference_fields.json').write_text(json.dumps({'pass':True,'scope':'12 already saved G1R12 native witnesses re-audited in Python; no new G1R13 solve claimed','results':rows},ensure_ascii=False,indent=2),encoding='utf-8')
