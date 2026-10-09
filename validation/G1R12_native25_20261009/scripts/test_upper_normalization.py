"""Controlled work-normalization experiments on saved native upper witnesses.

Not D03N replay. Does not alter a workbook, solver, tolerance or raw evidence.
"""
import json,math
import numpy as np
from prepare_inputs import ROOT
from case_locations import case_folder
from audit_batch import read_state
from independent_audit import audit,topology

def project(m,x,meta):
    before=audit(m,x,'upper',meta['fs'],meta['q'],meta['objective'],18*len(m.tri))
    work=before['work'][1]
    if not math.isfinite(work) or work<=0:raise ValueError('nonpositive reference work')
    field=x/work
    expected=before['value']/work
    after=audit(m,field,'upper',meta['fs'],meta['q'],expected,18*len(m.tri))
    return field,before,after

def run():
    records=[]
    for case in ['D01A','P02N']:
        folder=case_folder(case)
        status=json.loads((folder/'status.json').read_text(encoding='utf-8'))
        rec=json.loads((ROOT/'inputs'/f'{case}.json').read_text(encoding='utf-8'))
        m,x,meta=read_state(folder/'witness_upper.bin',rec['source_case']['policy'],'reference_total')
        baseline=audit(m,x,'upper',meta['fs'],meta['q'],meta['objective'],18*len(m.tri))
        assert baseline['pass'] and baseline['value']<=1+1e-8
        _,base_before,base_after=project(m,x,meta);assert base_after['pass']
        for deviation in [1.86894819176331e-7,-1.86894819176331e-7,
                          5.28104496799742e-7,-5.28104496799742e-7]:
            disturbed=x*((1+deviation)/baseline['work'][1])
            _,before,after=project(m,disturbed,meta)
            assert not before['pass'] and before['errors']['work']>1e-7
            assert after['pass'] and after['errors']['work']<=1e-7
            assert abs(after['value']-base_after['value'])<=2e-13
            records.append({'case':case,'controlled_work_deviation':deviation,
                            'before':before,'after':after,'feasible_field_recovered':True})
        for label,invalid in [('zero_work',np.zeros_like(x)),('negative_work',-x)]:
            try:project(m,invalid,meta)
            except ValueError:records.append({'case':case,'injection':label,'rejected':True})
            else:raise AssertionError(label)
        # Positive reference work does not justify bypassing the other audits.
        _,_,edges=topology(m)
        edge=next(e for e in edges if e.elements[1]<0 and e.kind=='fixed')
        invalid=x.copy();invalid[edge.elements[0]*12+edge.loc[0][0]*2]=.01
        _,before,after=project(m,invalid,meta)
        assert not after['pass'] and after['errors']['boundary']>1e-7
        records.append({'case':case,'injection':'fixed_boundary_velocity',
                        'before':before,'after':after,'rejected':True})
    out={'pass':True,'records':records,'actual_D03N_field_available':False,
         'actual_D03N_replayed':False,'vba_repair_implemented':False,
         'scope':'physical feasible-witness normalization only; no OPTIMAL or native root proof'}
    (ROOT/'underflow_diagnosis/upper_normalization_tests.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'pass':True,'native_witness_cases':['D01A','P02N'],'experiments':len(records),'actual_D03N_replayed':False}))

if __name__=='__main__':run()
