"""Independent original-q8 side audit; starts no solver."""
import argparse
import json
from pathlib import Path
import sys
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'python'))
import accuracy_driver as driver
import report_accuracy as report

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--case',required=True);p.add_argument('--mode',choices=['lower','upper'],required=True)
    p.add_argument('--folder',type=Path,required=True);p.add_argument('--output',type=Path,required=True);args=p.parse_args()
    _,rec=driver.load_model(args.case)
    meta=args.folder/('best_'+args.mode+'_certificate.json')
    candidate=report.collect_candidate(meta,args.mode,args.case,rec,[args.folder.resolve()],1e-4) if meta.exists() else None
    result=dict(case=args.case,mode=args.mode,target_met=False,two_percent_pair_evaluated=False,
        status='NO_SIDE_WITNESS' if candidate is None else 'VERIFIED_OWN_ROOT' if candidate['root_complete'] else 'VERIFIED_FIELD_ROOT_INCOMPLETE')
    if candidate:
        result.update(fs=candidate['fs'],elements=len(candidate['mesh'].tri),physical_pass=True,witness=True,
            root_complete=candidate['root_complete'],audit=candidate['audit'],source_sha256=candidate['source_hashes'])
    driver.write(args.output,result)
    print(json.dumps({k:result.get(k) for k in ('case','mode','status','fs','root_complete')}),flush=True)
    raise SystemExit(0 if candidate else 2)
