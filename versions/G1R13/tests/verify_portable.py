"""Replay saved native witnesses independently; no Excel/GitHub/original desktop packages required."""
from pathlib import Path
import sys,json,struct
import numpy as np
ROOT=Path(__file__).resolve().parents[1]
sys.path[:0]=[str(ROOT/'validation_runtime/python'),str(ROOT/'validation_runtime/vendor')]
from snapshot_reader import read_state
from audit_precision import audit_precision
from independent_audit import audit
from zero_phi import audit_zero
def vector(p):
    raw=p.read_bytes();assert raw[:8]==b'P06X0001';n=struct.unpack_from('<i',raw,8)[0];assert len(raw)==12+8*n
    return np.frombuffer(raw,dtype='<f8',offset=12).copy()
def run():
    rows=[];m,x,meta=read_state(ROOT/'evidence/D03N_original_rejected.mesh.bin','hold_initial','reference_total');m.provenance['case']='D03N'
    a=audit_precision(m,x,meta['fs'],'upper');assert not a['accepted'];rows.append({'case':'D03N','kind':'original_rejected','audit':a,'expected_reject':True})
    for name in ['original_normalized','solved']:
        p=ROOT/'single_point/D03N'/f'{name}.x.bin'
        if p.exists():
            a=audit_precision(m,vector(p),meta['fs'],'upper');assert a['accepted'];rows.append({'case':'D03N','kind':name,'audit':a})
    for case in ['D07A','D03N']:
        folder=ROOT/'native_retry'/case
        st=folder/'status.json'
        if not st.exists():continue
        rec=json.loads(st.read_text(encoding='utf-8'))
        if not rec.get('fs_obtained'):continue
        inp=json.loads((ROOT/'inputs'/f'{case}.json').read_text(encoding='utf-8'))
        for side in ['lower','upper']:
            m,x,meta=read_state(folder/f'witness_{side}.bin',inp['source_case']['policy'],'reference_total');m.provenance['case']=case
            a=audit_precision(m,x,meta['fs'],side)
            if side=='lower':
                precision_pass=max((v for k,v in a['raw_errors'].items() if k!='load_one'),default=0.)<=1e-7 and a['load_factor']>=1.-1e-8
                contract='lambda_max endpoint: lambda >= 1 - 1e-8; raw physical errors <= 1e-7'
            else:precision_pass=a['accepted'];contract='upper collapse / reference work'
            assert precision_pass,a
            d=audit(m,x,side,meta['fs'],meta['q'],meta['objective'],18*len(m.tri))
            assert side=='upper' or d['pass'],d
            assert meta['fs']==rec[side+'_fs']
            rows.append({'case':case,'kind':side+'_final','metadata':meta,'audit':a,'bound_contract':contract,'bound_contract_pass':precision_pass,'ordinary_double_audit':d})
    result={'pass':True,'scope':'stored binary64 fields, independent 50-digit physical integration/flow/edge audit; not rigorous interval proof','results':rows}
    print(json.dumps({'pass':True,'audited_fields':len(rows),'Fs_results_cases':[c for c in ['D07A','D03N'] if any(r['case']==c and r['kind']=='lower_final' for r in rows)]}))
    return result
if __name__=='__main__':
    result=run();(ROOT/'results/portable_audit.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
