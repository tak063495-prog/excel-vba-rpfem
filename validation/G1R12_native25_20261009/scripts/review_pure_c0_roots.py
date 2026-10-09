"""Record pure-c0 root contract evidence; no rounded identity proof or repair."""
import json,csv,collections
from prepare_inputs import ROOT,REPO
from case_locations import case_folder
from report_progress import collect

def run():
    rows=collect();assert all(r['status'] not in ['RUNNING','NOT_RUN','PREPARING'] for r in rows)
    records=[]
    for row in rows:
        rec=json.loads((ROOT/'inputs'/f'{row["case"]}.json').read_text(encoding='utf-8'))
        if any(m['c']!=0 for m in rec['source_case']['materials']):continue
        brackets=[];statuses=collections.Counter()
        for p in (case_folder(row['case'])/'RPFEM_logs').glob('*analysis.tsv'):
            with p.open(encoding='utf-16',newline='') as f:
                for event in csv.DictReader(f,delimiter='\t'):
                    if event['event']=='p4_final_bracket':brackets.append(event)
                    elif event['event']=='solve_end':
                        detail=dict(t.split('=',1) for t in event['detail'].split(';') if '=' in t)
                        statuses[detail.get('status','')]+=1
        records.append({'case':row['case'],'status':row['status'],'fs_obtained':row['fs_obtained'],
                        'root_flags':row.get('summary_fields',[])[12:16],
                        'target_unmet_reasons':row.get('target_unmet_reasons',[]),
                        'solve_status_counts':dict(statuses),'reported_final_brackets':brackets,
                        'stored_owner_endpoint_statuses_exported':False})
    source=REPO/'versions/G1R12/src/RPX_Fs.bas'
    lines=source.read_text(encoding='utf-8-sig').splitlines()
    guards=[{'line':i,'code':s.strip()} for i,s in enumerate(lines,1)
            if 'stored.statusA <> "OPTIMAL" Or stored.statusB <> "OPTIMAL"' in s]
    assert len(guards)==1
    result={'scope':'native logs + static contract review, not a solved root proof',
            'guard':guards[0],'cases':records,'vba_unchanged':True,
            'finding':'Pure-c0 equations are homogeneous in stress/load. Brackets can use native UNBOUNDED or INFEASIBLE certificates; the actual-root validity gate only accepts OPTIMAL endpoints.',
            'causality_limit':'Exact statuses and certificates of the saved best owner were not exported. Do not attribute every invalid root flag solely to this guard or certify a root by approximate log matching.',
            'required_repair_validation':'Store typed endpoint proof and exact model/q/load/Fs ownership. Verify both directional certificates and width before accepting a root; do not substitute OPTIMAL strings.'}
    (ROOT/'results/pure_c0_root_review.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'reviewed':len(records),'guard_line':guards[0]['line'],'vba_unchanged':True}))

if __name__=='__main__':run()
