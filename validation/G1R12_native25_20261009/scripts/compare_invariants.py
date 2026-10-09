"""Compare separately executed diagnostics; do not imply elastic sensitivity."""
import json,hashlib
from prepare_inputs import ROOT
from case_locations import case_folder
from report_progress import collect

def run():
    records={r['case']:r for r in collect()}
    base=records['P02A'];assert base.get('fs_obtained')
    rows=[]
    names={'D04A':'E/10 (inactive model parameter)','D09A':'nu=.49 (inactive model parameter)',
           'D06A':'translation (+1000,+2000)','D05A':'length x1000 / gamma divided by1000'}
    for case,name in names.items():
        r=records[case];assert r.get('fs_obtained')
        overlap=max(base['lower_fs'],r['lower_fs'])<=min(base['upper_fs'],r['upper_fs'])
        row={'case':case,'diagnostic':name,'independent_native_runs':True,'lower_fs':r['lower_fs'],'upper_fs':r['upper_fs'],
             'lower_relative_difference_from_P02A':(r['lower_fs']-base['lower_fs'])/base['lower_fs'],
             'upper_relative_difference_from_P02A':(r['upper_fs']-base['upper_fs'])/base['upper_fs'],
             'audited_intervals_overlap':overlap,'native_adopted_elements':r['element_counts'],
             'scope':'comparison of actual audited intervals; adapted meshes may differ, not a claim of identical discretization'}
        if case in ['D04A','D09A']:
            row['lower_payload_bytes_identical']= (case_folder(case)/'witness_lower.bin').read_bytes()==(case_folder('P02A')/'witness_lower.bin').read_bytes()
            row['upper_payload_bytes_identical']= (case_folder(case)/'witness_upper.bin').read_bytes()==(case_folder('P02A')/'witness_upper.bin').read_bytes()
            row['elastic_sensitivity_tested']=False
        rows.append(row)
    result={'base_case':'P02A','base_interval':[base['lower_fs'],base['upper_fs']],'results':rows}
    (ROOT/'results/invariants.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(result,ensure_ascii=False))

if __name__=='__main__':run()
