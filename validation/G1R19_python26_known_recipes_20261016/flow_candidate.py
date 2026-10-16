"""One same-size flow geometry candidate from an identified fresh field."""
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
from mesh_geometry import flow_metric_flip,interpolate_velocity
np=core.np


def sha(raw):return hashlib.sha256(raw).hexdigest()


def fresh_origin(prior,case):
    report=json.loads((prior/'summary.json').read_text(encoding='utf-8'))
    if report.get('case')!=case:raise ValueError('Fresh-run case differs')
    kind=report.get('execution_scope',{}).get('kind','')
    allowed=('cold_start_new_python_policy','fresh_automatic_candidate_then_bounded_generic_policy',
             'fresh_cold_start_with_new_generic_policy','fresh_pipeline_continuation')
    if kind in allowed:return report
    # A running continuation has not yet received the wrapper's final scope.
    # Its exact predecessor snapshot must itself be a new cold/fresh attempt.
    if kind=='saved_witness_reaudit_and_new_python_solves':
        path=prior/'predecessor_input_snapshot/summary.json'
        if path.is_file():
            previous=json.loads(path.read_text(encoding='utf-8'))
            if previous.get('case')==case and previous.get('execution_scope',{}).get('kind') in allowed:
                return report
    raise ValueError('Historical saved fields are excluded from this fresh flow comparison')


def geometry_moments(mesh):
    area=core.topology(mesh)[0]
    moments=area[:,None]*mesh.xy[mesh.tri].mean(axis=1)
    count=len(mesh.materials)
    return np.column_stack((np.bincount(mesh.material_id,weights=area,minlength=count),
        *[np.bincount(mesh.material_id,weights=moments[:,k],minlength=count) for k in range(2)]))


def run(args):
    started=time.perf_counter()
    if (args.output.exists() or not isinstance(args.passes,int) or args.passes<0 or
            not driver.finite(args.case_seconds,args.solve_seconds,args.initial_span,args.tolerance) or
            args.case_seconds<=0 or args.solve_seconds<=0 or args.initial_span<=0 or not 0<args.tolerance<=1e-4):
        raise ValueError('Fresh output, finite budgets/span and unchanged root tolerance required')
    source=args.seed.resolve();prior=args.prior_run.resolve()
    if not source.is_relative_to(prior):raise ValueError('Source must belong to identified fresh run')
    origin=fresh_origin(prior,args.case)
    budget=driver.CaseBudget(args.case_seconds,args.solve_seconds)
    frozen_sources=driver.source_hashes()
    frozen_sources[str(Path(__file__).resolve())]=sha(Path(__file__).read_bytes())
    frozen_input_paths=[ROOT/'INPUTS_26.json',driver.CATALOG/(args.case+'.json')]
    frozen_inputs={str(p.resolve()):sha(p.read_bytes()) for p in frozen_input_paths}
    names=[f'best_upper_{suffix}' for suffix in ('mesh.npz','field.npz','certificate.json')]
    raw={name:(source/name).read_bytes() for name in names}
    args.output.mkdir(parents=True)
    snapshot=args.output/'source_witness';snapshot.mkdir()
    for name,value in raw.items():(snapshot/name).write_bytes(value)
    if any((source/name).read_bytes()!=value or (snapshot/name).read_bytes()!=value for name,value in raw.items()):
        raise ValueError('Source changed during exact copy')
    driver.write(args.output/'fresh_origin_snapshot.json',origin)
    _,record=driver.load_model(args.case)
    if record.get('water',{}).get('coupling'):
        raise ValueError('This comparison does not transfer a coupled pressure discretization')
    parent=driver.load_saved_mesh(snapshot/names[0],record)
    with np.load(snapshot/names[1]) as data:
        if str(data['mode'])!='upper' or int(data['q'])!=8:raise ValueError('Upper q8 source required')
        field,fs=data['x'].copy(),float(data['fs'])
    metadata=json.loads(raw[names[2]].decode('utf-8'))
    identity=dict(case=args.case,mode='upper',q=8,fs=fs,mesh_id=parent.mesh_id,physical_model_id=parent.physical_model_id)
    if any(metadata.get(key)!=value for key,value in identity.items()):raise ValueError('Source identity differs')
    budget.check('source_original_audit')
    audit=core.evaluate(core.Prepared(parent,fs,'upper',8,reduction='none'),field)
    if not audit['physical_pass'] or not audit['witness']:raise ValueError('Source original q8 witness rejected')
    driver.write(args.output/'source_original_audit.json',audit)
    child,change=flow_metric_flip(parent,record,field,args.passes)
    before,after=geometry_moments(parent),geometry_moments(child)
    error=float(np.max(abs(before-after),initial=0))
    if not np.allclose(before,after,rtol=1e-12,atol=1e-10):raise ValueError('Material area/first moments changed')
    for key,old in parent.boundary.items():
        new=child.boundary[key]
        if old['rigid']!=new['rigid']:raise ValueError('Exterior rigid record changed')
    change['first_moment_guard']=dict(before=before.tolist(),after=after.tolist(),maximum_error=error)
    driver.save_mesh(args.output/'candidate_mesh.npz',child)
    driver.write(args.output/'mesh_change.json',change)
    budget.check('interpolation')
    transported=interpolate_velocity(parent,child,field,8)
    prepared=core.Prepared(child,fs,'upper',8,reduction='none')
    E,b,G,h,_,_=prepared.original
    rows=np.flatnonzero(b!=0)
    if len(rows)!=1 or np.any(h!=0):raise ValueError('Positive homogeneous upper normalization layout required')
    work=float((E[int(rows[0])]@transported)[0]/b[int(rows[0])])
    if not driver.finite(work) or work<=0:raise ValueError('Nonpositive/nonfinite interpolated work')
    transported/=work
    transferred_audit=core.evaluate(prepared,transported)
    driver.write(args.output/'interpolated_original_audit.json',transferred_audit)
    np.savez_compressed(args.output/'interpolated_candidate.npz',x=transported,fs=fs,mode='upper',q=8)
    warm=None
    if transferred_audit['witness']:
        warm=driver.Certificate('upper',fs,child,transported,transferred_audit,
            'full original-q8 audited nonnested interpolation; root not transferred',False)
    candidate=None;root=None
    if not args.prepare_only:
        delta=max(1e-4,args.initial_span*fs)
        candidate,root=driver.search(child,'upper',(max(1e-5,fs-delta),fs+delta),args.output/'root',
            args.case+'_fresh_flow',args.tolerance,args.solve_seconds,warm,
            budget=budget,bisection=True,max_trials=args.pilot_trials)
    current_sources=driver.source_hashes()
    current_sources[str(Path(__file__).resolve())]=sha(Path(__file__).read_bytes())
    unchanged=(current_sources==frozen_sources and
        all(sha(Path(path).read_bytes())==value for path,value in frozen_inputs.items()) and
        all((source/name).read_bytes()==value for name,value in raw.items()))
    if not unchanged:raise ValueError('Frozen implementation, input or source field changed')
    result=dict(case=args.case,scope='one same-element-count flow candidate from identified current-session field',
        mode='upper',q=8,two_percent_pair_evaluated=False,target_met=False,
        source_path=str(source),fresh_run_path=str(prior),source_identity=identity,
        source_original_audit=audit,source_sha256={name:sha(value) for name,value in raw.items()},
        source_root_transferred=False,old_source_mesh_retained=True,historical_field_imported=False,
        mesh_change=change,interpolated_original_audit=transferred_audit,
        interpolation_used_as_warm=warm is not None,root=root,
        seconds=time.perf_counter()-started,case_budget=budget.record(),
        source_unchanged=unchanged,source_sha256_before=frozen_sources,source_sha256_after=current_sources,
        frozen_input_sha256=frozen_inputs,optimizer_started=not args.prepare_only)
    if candidate is not None:
        result['upper']=candidate.publish(args.output)
        result['stronger_physical_upper']=candidate.fs<fs
        result['root_complete']=candidate.root_complete
    result['status']=('PREPARED_NO_SOLVER' if args.prepare_only else 'BUDGET_EXHAUSTED' if budget.remaining()<=0
        else 'UPPER_ROOT_COMPLETED' if candidate and candidate.root_complete else 'UPPER_FIELD_ONLY' if candidate else 'NO_UPPER_FIELD')
    driver.write(args.output/'summary.json',result)
    print('FLOW_CANDIDATE',args.case,result['status'],'elements',len(child.tri),'flips',change['flips'],
          'warm',warm is not None,'new_upper',candidate.fs if candidate else None,'seconds',result['seconds'],flush=True)
    return result


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--case',required=True);p.add_argument('--seed',type=Path,required=True)
    p.add_argument('--prior-run',type=Path,required=True);p.add_argument('--output',type=Path,required=True)
    p.add_argument('--passes',type=int,default=5);p.add_argument('--case-seconds',type=float,default=300.)
    p.add_argument('--solve-seconds',type=float,default=30.);p.add_argument('--initial-span',type=float,default=.005)
    p.add_argument('--tolerance',type=float,default=1e-4);p.add_argument('--pilot-trials',type=int)
    p.add_argument('--prepare-only',action='store_true');args=p.parse_args()
    report=run(args);raise SystemExit(0 if report['status'] in ('PREPARED_NO_SOLVER','UPPER_ROOT_COMPLETED') else 2)
