"""New own-mesh solves on geometry only; hints must come from this fresh session."""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import sys
import time
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'python'))
import accuracy_driver as driver
import experiment_core as core
np=core.np

def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def moments(mesh):
    area=core.topology(mesh)[0];n=len(mesh.materials)
    center=mesh.xy[mesh.tri].mean(axis=1)
    return np.column_stack([np.bincount(mesh.material_id,weights=area,minlength=n)]+
        [np.bincount(mesh.material_id,weights=area*center[:,k],minlength=n) for k in range(2)])

def read_hint(path,case,mode,model):
    data=json.loads(path.read_text(encoding='utf-8'));scope=data.get('execution_scope')
    allowed={'fresh_pipeline_continuation','cold_start_new_python_policy',
        'fresh_automatic_candidate_then_bounded_generic_policy','fresh_cold_start_with_new_generic_policy'}
    if not isinstance(scope,dict) or scope.get('kind') not in allowed or data.get('case')!=case:
        raise ValueError('Hint must be a current-session identified fresh case; historical results are excluded')
    bound=data.get(mode) or {}
    if bound.get('physical_model_id')!=model or bound.get('q')!=8 or bound.get('mode')!=mode:
        raise ValueError('Hint physical model, quadrature or mode differs')
    fs=bound.get('fs')
    if not driver.finite(fs) or fs<=0:raise ValueError('Nonpositive/nonfinite hint')
    return float(fs),dict(case=case,mode=mode,fs=fs,physical_model_id=model,
        source_scope=scope['kind'],source_status=data.get('status'),source_root_imported=False,
        source_field_imported=False,purpose='Initial search location only; fresh same-own-mesh witnesses/signs required')

def run(args):
    started=time.perf_counter()
    if args.output.exists() or not driver.finite(args.initial_span) or args.initial_span<=0:raise ValueError('Fresh output and positive span required')
    budget=driver.CaseBudget(args.case_seconds,args.solve_seconds)
    sources=driver.source_hashes();sources[str(Path(__file__).resolve())]=sha(__file__)
    paths=[ROOT/'INPUTS_26.json',driver.CATALOG/(args.case+'.json'),args.geometry,args.hint_summary]
    inputs={str(p.resolve()):sha(p) for p in paths}
    coarse,record=driver.load_model(args.case)
    if record.get('water',{}).get('coupling'):raise ValueError('No coupled pressure geometry replay')
    mesh=driver.load_saved_mesh(args.geometry,record)
    before,after=moments(coarse),moments(mesh)
    if not np.allclose(before,after,rtol=1e-12,atol=1e-10):raise ValueError('Material area or first moments differ from canonical model')
    hint,hint_info=read_hint(args.hint_summary,args.case,args.mode,mesh.physical_model_id)
    span=args.initial_span*max(1.,hint)
    initial=(max(1e-5,hint-span),hint+span)
    args.output.mkdir(parents=True);driver.save_mesh(args.output/'candidate_mesh.npz',mesh)
    driver.write(args.output/'fresh_hint_snapshot.json',json.loads(args.hint_summary.read_text(encoding='utf-8')))
    info=dict(case=args.case,mode=args.mode,q=8,execution_scope=dict(kind='known_recipe_geometry_replay_new_solves',
        historical_geometry_only=True,historical_field_imported=False,historical_fs_imported=False),
        geometry_path=str(args.geometry.resolve()),geometry_sha256=sha(args.geometry),
        hint=hint_info,initial_bracket=list(initial),material_moments=dict(canonical=before.tolist(),candidate=after.tolist()),
        elements=len(mesh.tri),mesh_id=mesh.mesh_id,physical_model_id=mesh.physical_model_id,
        source_sha256_before=sources,input_sha256=inputs,optimizer_started=False,status='PREPARED_NO_SOLVER',target_met=False)
    driver.write(args.output/'summary.json',info)
    print('PREPARED_GEOMETRY',args.case,args.mode,len(mesh.tri),'initial',initial,flush=True)
    if args.prepare_only:return info
    candidate,root=driver.search(mesh,args.mode,initial,args.output/'root',args.case+'_new_'+args.mode+'_known_geometry',
        1e-4,args.solve_seconds,budget=budget,bisection=True)
    current=driver.source_hashes();current[str(Path(__file__).resolve())]=sha(__file__)
    unchanged=current==sources and all(sha(p)==h for p,h in inputs.items())
    if not unchanged:raise ValueError('Frozen source, geometry or current-session hint changed')
    info.update(optimizer_started=True,root=root,case_budget=budget.record(),seconds=time.perf_counter()-started,
        source_unchanged=unchanged,source_sha256_after=current)
    if candidate is not None:info[args.mode]=candidate.publish(args.output);info['root_complete']=candidate.root_complete
    info['status']=('BUDGET_EXHAUSTED' if budget.remaining()<=0 else 'OWN_ROOT_COMPLETED' if candidate and candidate.root_complete
        else 'PHYSICAL_FIELD_ROOT_INCOMPLETE' if candidate else 'NO_PHYSICAL_FIELD')
    driver.write(args.output/'summary.json',info)
    print('GEOMETRY_DONE',args.case,args.mode,info['status'],candidate.fs if candidate else None,'seconds',info['seconds'],flush=True)
    return info

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--case',required=True);p.add_argument('--mode',choices=['lower','upper'],required=True)
    p.add_argument('--geometry',type=Path,required=True);p.add_argument('--hint-summary',type=Path,required=True)
    p.add_argument('--output',type=Path,required=True);p.add_argument('--initial-span',type=float,default=.005)
    p.add_argument('--case-seconds',type=float,default=450.);p.add_argument('--solve-seconds',type=float,default=30.)
    p.add_argument('--prepare-only',action='store_true');args=p.parse_args();result=run(args)
    raise SystemExit(0 if result['status'] in ('PREPARED_NO_SOLVER','OWN_ROOT_COMPLETED') else 2)
