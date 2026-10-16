"""Fresh lower solve on a regenerated historical geometry recipe, without fields/Fs."""
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

def run(args):
    started=time.perf_counter()
    if args.output.exists() or args.rounds<0 or args.max_elements<1:raise ValueError('Fresh output and bounded refinement required')
    budget=driver.CaseBudget(args.case_seconds,args.solve_seconds)
    sources=driver.source_hashes();sources[str(Path(__file__).resolve())]=sha(__file__)
    paths=[ROOT/'INPUTS_26.json',driver.CATALOG/(args.case+'.json'),args.geometry,args.expected_geometry]
    inputs={str(p.resolve()):sha(p) for p in paths}
    args.output.mkdir(parents=True)
    _,record=driver.load_model(args.case)
    mesh=driver.load_saved_mesh(args.geometry,record)
    if record.get('water',{}).get('coupling'):raise ValueError('Dry geometry recipe only')
    recipe=[dict(method='input_geometry_only',elements=len(mesh.tri),mesh_id=mesh.mesh_id)]
    for _ in range(args.rounds):
        mesh=core.uniform_refine(mesh,record)
        if len(mesh.tri)>args.max_elements:raise ValueError('Element cap')
        recipe.append(dict(method='uniform_red4',elements=len(mesh.tri),mesh_id=mesh.mesh_id))
    with np.load(args.expected_geometry,allow_pickle=False) as saved:
        exact=all(np.array_equal(saved[k],getattr(mesh,k)) for k in ('xy','tri','material_id'))
    if not exact:raise ValueError('Regenerated mesh differs from historical recipe geometry')
    driver.save_mesh(args.output/'candidate_mesh.npz',mesh)
    info=dict(case=args.case,execution_scope='known_recipe_geometry_replay_new_solves',
        historical_field_imported=False,historical_fs_imported=False,recipe=recipe,
        regenerated_geometry_exact_match=True,physical_model_id=mesh.physical_model_id,
        initial_bracket=[.4,1.3],q=8,target_gap=.02,source_sha256=sources,input_sha256=inputs,
        status='PREPARED_NO_SOLVER',optimizer_started=False,target_met=False)
    driver.write(args.output/'summary.json',info)
    print('PREPARED_LOWER',args.case,len(mesh.tri),'exact',exact,'seconds',time.perf_counter()-started,flush=True)
    if args.prepare_only:return info
    candidate,root=driver.search(mesh,'lower',(.4,1.3),args.output/'root',args.case+'_fresh_uniform_recipe',
        1e-4,args.solve_seconds,budget=budget,bisection=True)
    current=driver.source_hashes();current[str(Path(__file__).resolve())]=sha(__file__)
    unchanged=current==sources and all(sha(p)==h for p,h in inputs.items())
    if not unchanged:raise ValueError('Frozen sources or inputs changed')
    info.update(optimizer_started=True,root=root,case_budget=budget.record(),
        seconds=time.perf_counter()-started,source_unchanged=unchanged,source_sha256_after=current)
    if candidate is not None:
        audit=core.evaluate(core.Prepared(mesh,candidate.fs,'lower',8,reduction='none'),candidate.field)
        if not audit['physical_pass'] or not audit['witness']:raise ValueError('Final original physical audit rejected')
        info['independent_same_process_original_audit']=audit
        info['lower']=candidate.publish(args.output)
        info['root_complete']=candidate.root_complete
    info['status']=('BUDGET_EXHAUSTED' if budget.remaining()<=0 else 'LOWER_ROOT_COMPLETED' if candidate and candidate.root_complete
                    else 'LOWER_FIELD_ONLY' if candidate else 'NO_LOWER_FIELD')
    driver.write(args.output/'summary.json',info)
    print('LOWER_DONE',args.case,info['status'],candidate.fs if candidate else None,'seconds',info['seconds'],flush=True)
    return info

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--case',required=True);p.add_argument('--geometry',type=Path,required=True)
    p.add_argument('--expected-geometry',type=Path,required=True);p.add_argument('--output',type=Path,required=True)
    p.add_argument('--rounds',type=int,default=1);p.add_argument('--max-elements',type=int,default=4000)
    p.add_argument('--case-seconds',type=float,default=450.);p.add_argument('--solve-seconds',type=float,default=30.)
    p.add_argument('--prepare-only',action='store_true');args=p.parse_args()
    result=run(args);raise SystemExit(0 if result['status'] in ('PREPARED_NO_SOLVER','LOWER_ROOT_COMPLETED') else 2)
