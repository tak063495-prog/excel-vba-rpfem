"""Finish a current-session own-mesh root without transporting a proof."""
from pathlib import Path
import argparse
import hashlib
import json
import sys
import time
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'python'))
import accuracy_driver as driver
import report_accuracy as report

def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def run(args):
    started=time.perf_counter()
    if args.output.exists():raise ValueError('Fresh output required')
    budget=driver.CaseBudget(args.case_seconds,args.solve_seconds)
    source=args.source.resolve();origin=source/'summary.json'
    document=json.loads(origin.read_text(encoding='utf-8'));scope=document.get('execution_scope')
    kind=scope.get('kind') if isinstance(scope,dict) else scope
    if document.get('case')!=args.case or kind not in ('known_recipe_geometry_replay_new_solves','fresh_pipeline_continuation','current_session_side_root_continuation'):
        raise ValueError('Only identified current-session freshly solved fields may continue')
    sources=driver.source_hashes();sources[str(Path(__file__).resolve())]=sha(__file__)
    paths=[origin,ROOT/'INPUTS_26.json',driver.CATALOG/(args.case+'.json')]
    _,rec=driver.load_model(args.case)
    candidate=report.collect_candidate(source/('best_'+args.mode+'_certificate.json'),args.mode,args.case,rec,[source],1e-4)
    if candidate is None:raise ValueError('No unchanged original physical witness')
    paths.extend(candidate['tracked']);inputs={str(p.resolve()):sha(p) for p in paths}
    mesh=candidate['mesh'];fs=candidate['fs'];audit=candidate['audit']
    root=audit.get('discrete_root_search')or{}
    initial=root.get('bracket')
    valid=bool(isinstance(initial,list) and len(initial)==2 and driver.finite(*initial) and 0<initial[0]<initial[1]
        and root.get('mesh_id')==mesh.mesh_id and root.get('physical_model_id')==mesh.physical_model_id
        and root.get('q')==8 and root.get('mode')==args.mode and root.get('case')==args.case
        and root.get('policy')==mesh.provenance['psi_policy'])
    old_initial=list(initial) if valid else None
    if not valid:initial=[max(1e-5,fs-.005),fs+.005]
    elif initial[1]-initial[0]>1e-4:
        # A proposal only. The two new endpoints need fresh original signs;
        # a missed root triggers the unchanged ordinary bracket expansion.
        span=initial[1]-initial[0]
        initial=[initial[0]+.3*span,initial[1]-.3*span]
    warm=driver.Certificate(args.mode,fs,mesh,candidate['field'],audit,'Current-session source independently original-q8 reaudited',False)
    args.output.mkdir(parents=True)
    driver.write(args.output/'fresh_source_snapshot.json',document)
    initial_metadata=warm.publish(args.output)
    result=dict(case=args.case,mode=args.mode,q=8,execution_scope=dict(kind='current_session_side_root_continuation',
        historical_field_imported=False,historical_fs_imported=False),source=str(source),source_root_imported=False,
        initial_bracket=initial,source_bracket=old_initial,proposal='central_40_percent_fresh_signed_probe',input_sha256=inputs,source_sha256_before=sources,target_met=False,
        status='ROOT_CLOSURE_RUNNING',certificates={args.mode:initial_metadata})
    driver.write(args.output/'summary.json',result)
    best,newroot=driver.search(mesh,args.mode,initial,args.output/'root',args.case+'_close_current_'+args.mode,
        1e-4,args.solve_seconds,initial_witness=warm,budget=budget,bisection=True)
    current=driver.source_hashes();current[str(Path(__file__).resolve())]=sha(__file__)
    unchanged=current==sources and all(sha(p)==h for p,h in inputs.items())
    if not unchanged:raise ValueError('Source or input changed')
    result.update(root=newroot,seconds=time.perf_counter()-started,case_budget=budget.record(),source_unchanged=True,source_sha256_after=current)
    best=best or warm;result['certificates']={args.mode:best.publish(args.output)};result[args.mode]=result['certificates'][args.mode]
    result['status']='OWN_ROOT_COMPLETED' if best.root_complete else 'BUDGET_EXHAUSTED' if budget.remaining()<=0 else 'PHYSICAL_FIELD_ROOT_INCOMPLETE'
    result['root_complete']=best.root_complete
    driver.write(args.output/'summary.json',result)
    print('CLOSED_CURRENT',args.case,args.mode,best.fs,best.root_complete,result['seconds'],flush=True)
    return result

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--case',required=True);p.add_argument('--mode',choices=['lower','upper'],required=True)
    p.add_argument('--source',type=Path,required=True);p.add_argument('--output',type=Path,required=True)
    p.add_argument('--case-seconds',type=float,default=120.);p.add_argument('--solve-seconds',type=float,default=30.)
    result=run(p.parse_args());raise SystemExit(0 if result['root_complete'] else 2)
