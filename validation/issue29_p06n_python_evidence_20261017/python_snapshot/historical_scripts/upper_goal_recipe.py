"""Reproduce a geometry recipe with new LP/QP primals and original dual guides."""
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
import report_accuracy as collector
import experiment_core as core
from original_dual import solve_original_dual,audit_original_dual
from upper_proximal_candidate import solve_candidate
np=core.np

def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def moments(mesh):
    area=core.topology(mesh)[0];center=mesh.xy[mesh.tri].mean(axis=1);n=len(mesh.materials)
    return np.column_stack([np.bincount(mesh.material_id,weights=area,minlength=n)]+
        [np.bincount(mesh.material_id,weights=area*center[:,k],minlength=n) for k in range(2)])

def normalized_candidate(prepared,x):
    x=np.asarray(x,dtype=float).copy();before=core.evaluate(prepared,x)
    E,b,G,h,dims,cost=prepared.original;rows=np.flatnonzero(b!=0)
    work=None
    if np.isfinite(x).all() and len(rows)==1 and np.all(h==0):
        work=float((E[int(rows[0])]@x)[0]/b[int(rows[0])])
        if np.isfinite(work) and work>0:x/=work
    after=core.evaluate(prepared,x)
    return x,after,dict(method='positive homogeneous original work normalization',work_ratio=work,
        before_original_audit=before,after_original_audit=after)

def run(args):
    started=time.perf_counter()
    if (args.output.exists() or not driver.finite(args.target_gap,args.schedule_gap,args.alpha,args.qp_seconds)
        or not 0<args.schedule_gap<args.target_gap<=.02 or args.max_steps<1 or args.alpha<=0 or args.qp_seconds<=0):
        raise ValueError('Fresh output, bounded strict goal schedule and positive QP settings required')
    budget=driver.CaseBudget(args.case_seconds,max(args.solve_seconds,args.qp_seconds))
    sources=driver.source_hashes();sources[str(Path(__file__).resolve())]=sha(__file__)
    lower_folder=args.lower_source.resolve();origin=lower_folder/'summary.json'
    document=json.loads(origin.read_text(encoding='utf-8'));scope=document.get('execution_scope')
    kind=scope.get('kind') if isinstance(scope,dict) else scope
    if document.get('case')!=args.case or kind not in ('known_recipe_geometry_replay_new_solves','current_session_side_root_continuation'):
        raise ValueError('Only a current-session newly solved lower may set the goal')
    coarse,record=driver.load_model(args.case)
    lower=collector.collect_candidate(lower_folder/'best_lower_certificate.json','lower',args.case,record,[lower_folder],1e-4)
    if lower is None or not lower['root_complete']:raise ValueError('A freshly original-audited completed lower is required')
    mesh=driver.load_saved_mesh(args.geometry,record)
    if record.get('water',{}).get('coupling') or not np.allclose(moments(coarse),moments(mesh),rtol=1e-12,atol=1e-10):
        raise ValueError('The canonical dry material domain must be unchanged')
    if mesh.physical_model_id!=lower['mesh'].physical_model_id:raise ValueError('Model mismatch')
    paths=[ROOT/'INPUTS_26.json',driver.CATALOG/(args.case+'.json'),args.geometry,origin]+list(lower['tracked'])
    inputs={str(p.resolve()):sha(p) for p in paths}
    goal=lower['fs']*(2+args.schedule_gap)/(2-args.schedule_gap)
    args.output.mkdir(parents=True)
    driver.save_mesh(args.output/'candidate_mesh.npz',mesh)
    driver.write(args.output/'fresh_lower_snapshot.json',document)
    info=dict(case=args.case,execution_scope=dict(kind='known_recipe_geometry_replay_new_solves',historical_geometry_imported=True,
        historical_field_imported=False,historical_fs_imported=False),target_gap=args.target_gap,schedule_gap=args.schedule_gap,
        solver_call_limits=dict(original_lp=args.solve_seconds,proximal_qp=args.qp_seconds,original_dual=30.),initial_goal_fs=goal,goal_formula='fresh_lower_fs * (2 + schedule_gap) / (2 - schedule_gap)',
        lower_source=str(lower_folder),lower_fs=lower['fs'],geometry_path=str(args.geometry.resolve()),
        source_sha256_before=sources,input_sha256=inputs,trials=[],target_met=False,root_complete=False,status='RUNNING')
    best=None;current_fs=goal
    def checkpoint():
        info.update(seconds=time.perf_counter()-started,case_budget=budget.record())
        if best is not None:info['upper']=best.publish(args.output)
        driver.write(args.output/'summary.json',info)
    def guide_from_vector(prepared,z,label):
        audit=audit_original_dual(prepared,z)
        np.savez_compressed(args.output/(label+'.npz'),dual=z,fs=prepared.fs,mode='upper',q=8,
            case=args.case,mesh_id=mesh.mesh_id,physical_model_id=mesh.physical_model_id,policy=mesh.provenance['psi_policy'])
        driver.write(args.output/(label+'_audit.json'),audit)
        if not audit['accepted']:return None
        return dict(case=args.case,mode='upper',q=8,fs=prepared.fs,mesh_id=mesh.mesh_id,
            physical_model_id=mesh.physical_model_id,policy=mesh.provenance['psi_policy'],dual=z)
    def close(guide):
        nonlocal best
        best,root=driver.search(mesh,'upper',(guide['fs'],best.fs),args.output/'closed_root',args.case+'_upper_goal_closure',
            1e-4,args.solve_seconds,initial_witness=best,initial_dual_guides=[guide],budget=budget,bisection=True)
        info['root']=root
    checkpoint()
    try:
        for step in range(args.max_steps):
            budget.check('original_upper_goal_prepare')
            prepared=core.Prepared(mesh,current_fs,'upper',8,reduction='none')
            if best is not None:
                evidence=solve_original_dual(prepared,seconds=budget.allowance('original_dual_before_primal',30.),return_dual=True)
                z=evidence.pop('dual',None);driver.write(args.output/f'step{step}_original_dual_helper.json',evidence)
                if z is not None:
                    guide=guide_from_vector(prepared,z,f'step{step}_original_dual')
                    if guide is not None:close(guide);checkpoint();break
            # A fresh original LP supplies an original-layout center, even when
            # its primal is not yet admissible. No archived velocity is read.
            reduced=core.Prepared(mesh,current_fs,'upper',8,reduction='affine')
            raw_path=args.output/f'step{step}_original_lp.npz'
            raw,center,_=core.solve(reduced,dict(driver.ACCURATE,time_limit=budget.allowance('original_lp',args.solve_seconds)),max_iter=350,
                primal_polish=True,save=raw_path)
            x,audit,norm=normalized_candidate(prepared,center)
            info['trials'].append(dict(step=step,fs=current_fs,method='original_lp',audit=audit,normalization=norm,solver_metadata=raw))
            guide=None
            if not audit['witness']:
                with np.load(raw_path,allow_pickle=False) as data:
                    z=data['z'].copy()
                guide=guide_from_vector(prepared,z,f'step{step}_original_lp_dual')
                if guide is not None:
                    if best is not None:close(guide)
                    else:info['status']='ORIGINAL_POSITIVE_AT_TWO_PERCENT_GOAL'
                    checkpoint();break
            if not audit['witness']:
                center=x
                for repeat in range(2):
                    x,modified_dual,qp=solve_candidate(prepared,center,args.alpha,
                        budget.allowance('original_cap_qp',args.qp_seconds))
                    np.savez_compressed(args.output/f'step{step}_qp{repeat}_raw.npz',x=x,fs=current_fs,mode='upper',q=8,
                        modified_auxiliary_dual_not_original_guide=modified_dual)
                    x,audit,norm=normalized_candidate(prepared,x)
                    info['trials'].append(dict(step=step,repeat=repeat,fs=current_fs,method='original_constraints_capped_proximal_qp',
                        audit=audit,normalization=norm,solver_metadata=qp,modified_dual_used_as_original_guide=False))
                    if audit['witness']:break
                    center=x.copy();checkpoint()
            if not audit['physical_pass'] or not audit['witness']:
                info['status']='NO_ORIGINAL_PHYSICAL_UPPER_AT_TRIAL';checkpoint();break
            best=driver.Certificate('upper',current_fs,mesh,x,audit,'Fresh original LP/QP; unchanged full original-q8 audit',False)
            checkpoint();print('GOAL_UPPER',args.case,step,current_fs,audit['value'],flush=True)
            current_fs=best.fs-.95e-4*max(1.,best.fs)
    except driver.BudgetExceeded as error:
        info.update(status='BUDGET_EXHAUSTED',budget_error=str(error))
    current=driver.source_hashes();current[str(Path(__file__).resolve())]=sha(__file__)
    unchanged=current==sources and all(sha(p)==h for p,h in inputs.items())
    if not unchanged:raise ValueError('Frozen source, original input, geometry or lower witness changed')
    info.update(source_unchanged=True,source_sha256_after=current)
    if best is not None:
        gap=2*(best.fs-lower['fs'])/(best.fs+lower['fs'])
        info.update(relative_width_symmetric=gap,root_complete=best.root_complete,
            target_met=bool(best.root_complete and lower['root_complete'] and 0<=gap<=args.target_gap))
        if info['target_met']:info['status']='VERIFIED_TARGET'
        elif info['status']=='RUNNING':info['status']='PHYSICAL_UPPER_ROOT_INCOMPLETE'
    elif info['status']=='RUNNING':info['status']='NO_UPPER_FIELD'
    checkpoint();print('GOAL_RECIPE_DONE',args.case,info['status'],info.get('relative_width_symmetric'),info['seconds'],flush=True)
    return info

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--case',required=True);p.add_argument('--geometry',type=Path,required=True)
    p.add_argument('--lower-source',type=Path,required=True);p.add_argument('--output',type=Path,required=True)
    p.add_argument('--target-gap',type=float,default=.02);p.add_argument('--schedule-gap',type=float,default=.019999)
    p.add_argument('--case-seconds',type=float,default=600.);p.add_argument('--solve-seconds',type=float,default=30.)
    p.add_argument('--qp-seconds',type=float,default=90.);p.add_argument('--alpha',type=float,default=.01)
    p.add_argument('--max-steps',type=int,default=4)
    result=run(p.parse_args());raise SystemExit(0 if result['target_met'] else 2)
