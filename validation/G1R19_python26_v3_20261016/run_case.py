"""Fresh bounded policy with exact upper velocity-unit conditioning.

The existing physical model, q8 gates and own-mesh root requirements are
unchanged. Geometry chooses a candidate; neither its formula nor a solver
status supplies a bound. No historical field or Fs is used by the cold path.
"""
from __future__ import annotations
import argparse
import copy
import hashlib
import json
from pathlib import Path
import platform
import sys
import time

ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'python'))
import accuracy_driver as driver
import c0_slope_wedge_driver as wedge
import experiment_core as core
from report_accuracy import root_is_complete

_ORIGINAL_SOLVE=driver.solve
_UPPER_DIAGNOSTICS=None
_UPPER_SEQUENCE=0


def program_digest(program):
    result=hashlib.sha256()
    for value in program:
        arrays=(value.indptr,value.indices,value.data) if core.sparse.issparse(value) else (core.np.asarray(value),)
        for array in arrays:
            result.update(str(array.shape).encode('ascii'))
            result.update(str(array.dtype).encode('ascii'))
            result.update(array.tobytes())
    return result.hexdigest()


def normalize_upper_work(prepared,source):
    """Reconstruct only existing velocity auxiliaries and audit positive scaling."""
    E,b,G,h,dims,cost=prepared.original
    nonzero=core.np.flatnonzero(b)
    evidence={'attempted':False,'accepted':False}
    if len(nonzero)!=1 or b[nonzero[0]]<=0 or core.np.any(h!=0):
        return None,dict(evidence,reason='Single positive reference-work and homogeneous original cones required')
    candidate=core.complete_upper_velocity(prepared.m,source,prepared.q)
    if candidate.shape!=source.shape or not core.np.isfinite(candidate).all():
        return None,dict(evidence,reason='Existing velocity completion changed the original layout')
    row=int(nonzero[0])
    work=float((E.getrow(row)@candidate)[0]/b[row])
    if not driver.finite(work) or work<=0:
        return None,dict(evidence,reason='No finite positive reference work')
    candidate=candidate/work
    audit=core.evaluate(prepared,candidate)
    evidence.update(attempted=True,work_before=work,positive_scale=1/work,
        audit=audit,accepted=bool(audit.get('physical_pass')),
        scope='same original coordinates, strength, loads and quadrature; primal candidate only')
    return candidate,evidence


def scaled_upper_solve(prepared,overrides=None,max_iter=250,duals=True,save=None,**kwargs):
    """Solve the identical original LP in velocity units x=s*u, not a new model.

    q_new=s*q, rhs_new=rhs/s, A_new=[E;-R*G] with a positive common multiplier
    per original cone. Restore x=s*u, y=y_new/s, z=R*z_new/s. Original vectors
    alone are re-audited; a solver status never supplies a sign or witness.
    """
    global _UPPER_SEQUENCE
    if prepared.mode!='upper':
        return _ORIGINAL_SOLVE(prepared,overrides,max_iter,duals,save,**kwargs)
    started=time.perf_counter()
    E,b,G,h,dims,cost=prepared.original
    work_rows=core.np.flatnonzero(b)
    if len(work_rows)!=1 or b[work_rows[0]]<=0 or not core.np.isfinite(b).all():
        return _ORIGINAL_SOLVE(prepared,overrides,max_iter,duals,save,**kwargs)
    scale=float(b[work_rows[0]])
    if not driver.finite(scale) or scale<=0:
        raise ValueError('A finite positive exact velocity coordinate scale is required')
    before=program_digest(prepared.original)
    row_scale=core.cone_row_scales(G,dims)
    A=core.sparse.vstack([E,-core.sparse.diags(row_scale)@G],format='csc')
    rhs=core.np.r_[b,h*row_scale]/scale
    objective=cost*scale
    if not all(core.np.isfinite(v).all() for v in (A.data,rhs,objective,row_scale)):
        raise ValueError('Nonfinite scaled original program')
    cones=[core.clarabel.ZeroConeT(len(b))]+[
        core.clarabel.NonnegativeConeT(d) if d==1 else core.clarabel.SecondOrderConeT(d) for d in dims]
    config=dict(overrides or {})
    # Restoring an original dual divides its residual by scale. The strict
    # existing profile avoids wasting attempts on a scaled-space-only pass.
    for key in ('tol_gap_abs','tol_gap_rel','tol_feas'):
        config[key]=min(float(config.get(key,1e-10)),1e-13)
    config['iterative_refinement_max_iter']=max(int(config.get('iterative_refinement_max_iter',0)),30)
    config['iterative_refinement_abstol']=min(float(config.get('iterative_refinement_abstol',1e-12)),1e-15)
    config['iterative_refinement_reltol']=min(float(config.get('iterative_refinement_reltol',1e-12)),1e-15)
    settings=core.settings(config,max_iter)
    solver=core.clarabel.DefaultSolver(core.sparse.csc_matrix((len(cost),len(cost))),objective,A,rhs,cones,settings)
    build=time.perf_counter()-started
    solution=solver.solve()
    field=scale*core.np.asarray(solution.x)
    coordinate_dual=core.np.asarray(solution.z)
    original_dual=core.np.r_[coordinate_dual[:len(b)],row_scale*coordinate_dual[len(b):]]/scale
    result=core.evaluate(prepared,field,original_dual if duals else None)
    raw_audit=copy.deepcopy(result)
    recovered,recovery=normalize_upper_work(prepared,field)
    if recovered is not None:
        recovered_result=core.evaluate(prepared,recovered,original_dual if duals else None)
        if driver.physical_attempt_rank(recovered_result,'upper')>driver.physical_attempt_rank(result,'upper'):
            result,field=recovered_result,recovered
            recovery['selected']=True
        else:recovery['selected']=False
    after=program_digest(prepared.original)
    if after!=before:raise RuntimeError('Original program changed during coordinate solve')
    result.update(case=prepared.m.provenance['case'],mode='upper',q=prepared.q,fs=prepared.fs,
        elements=len(prepared.m.tri),status=str(solution.status),iterations=solution.iterations,
        solver_seconds=float(solution.solve_time),build_seconds=build,
        end_to_end_seconds=time.perf_counter()-started,settings_overrides=config,
        max_iter=max_iter,reduced_numerical_gate=False,
        dual_restore={'method':'exact_original_full_space_positive_coordinate_restore'},
        original_program_sha256=before,original_program_unchanged=True,
        velocity_coordinate_scale=scale,common_positive_cone_row_scaling=True,
        raw_before_work_normalization=raw_audit,upper_work_recovery=recovery,
        provenance='same original full LP; exact positive scalar velocity units; original primal and original dual gates')
    _UPPER_SEQUENCE+=1
    if _UPPER_DIAGNOSTICS is not None:
        _UPPER_DIAGNOSTICS.mkdir(parents=True,exist_ok=True)
        stem=_UPPER_DIAGNOSTICS/f'upper_attempt_{_UPPER_SEQUENCE:04d}'
        core.np.savez_compressed(stem.with_suffix('.npz'),x=field,original_dual=original_dual,
            coordinate_x=core.np.asarray(solution.x),coordinate_dual=coordinate_dual,
            fs=prepared.fs,mode='upper',q=prepared.q,velocity_scale=scale,cone_row_scales=row_scale)
        driver.write(stem.with_suffix('.json'),result)
    if save:
        core.np.savez_compressed(save,x=field,z=original_dual,fs=prepared.fs,mode='upper',q=prepared.q)
    return result,field,original_dual


def implementation_hashes():
    hashes=driver.source_hashes()
    for path in (Path(__file__).resolve(),ROOT/'INPUTS_26.json'):
        hashes[path.relative_to(driver.REPO).as_posix()]=hashlib.sha256(path.read_bytes()).hexdigest()
    return hashes


def checked_certificate(certificate,case):
    """A fresh full original audit and strict root identity precede publication."""
    if certificate is None:return None
    mesh=certificate.mesh
    if mesh.provenance['case']!=case or not driver.finite(certificate.fs) or certificate.fs<=0:
        raise ValueError('Candidate case or Fs differs')
    audit=driver.evaluate(driver.Prepared(mesh,certificate.fs,certificate.mode,8,reduction='none'),certificate.field)
    if not audit.get('physical_pass') or not audit.get('witness'):
        raise ValueError('Fresh full original q8 witness gate rejected the candidate')
    root=certificate.result.get('discrete_root_search',{})
    if root:audit['discrete_root_search']=root
    metadata={'case':case,'mode':certificate.mode,'q':8,'fs':certificate.fs,
        'mesh_id':mesh.mesh_id,'physical_model_id':mesh.physical_model_id,
        'root_complete':certificate.root_complete,'audit':audit}
    complete=root_is_complete(metadata,mesh,certificate.fs,certificate.mode,case)
    return driver.Certificate(certificate.mode,certificate.fs,mesh,certificate.field.copy(),audit,
                              certificate.source,complete)


def read_certificate(folder,mode,record,case):
    path=folder/f'best_{mode}_certificate.json'
    if not path.is_file():return None
    metadata=json.loads(path.read_text(encoding='utf8'))
    # This reads only a subrun freshly generated by this invocation.
    mesh=driver.load_saved_mesh(folder/f'best_{mode}_mesh.npz',record)
    with driver.np.load(folder/f'best_{mode}_field.npz',allow_pickle=False) as data:
        if str(data['mode'])!=mode or int(data['q'])!=8:
            raise ValueError('Fresh subrun field mode or quadrature differs')
        field=data['x'].copy();fs=float(data['fs'])
    if (metadata.get('case')!=case or metadata.get('mode')!=mode or metadata.get('q')!=8 or
            metadata.get('fs')!=fs or metadata.get('mesh_id')!=mesh.mesh_id or
            metadata.get('physical_model_id')!=mesh.physical_model_id):
        raise ValueError('Fresh subrun metadata and field identities differ')
    certificate=driver.Certificate(mode,fs,mesh,field,metadata.get('audit',{}),
        str(path),metadata.get('root_complete') is True)
    return checked_certificate(certificate,case)


def retain(best,certificate):
    if certificate is None:return
    old=best.get(certificate.mode)
    if old is not None and old.mesh.physical_model_id!=certificate.mesh.physical_model_id:
        raise ValueError('Independent candidates changed the physical model')
    quality=certificate.fs if certificate.mode=='lower' else -certificate.fs
    previous=old.fs if old is not None and old.mode=='lower' else -old.fs if old is not None else 0.
    # Stronger unfinished fields remain in the subrun. Terminal output first
    # retains a complete own-root certificate, then improves its endpoint.
    if old is None or (certificate.root_complete,quality)>(old.root_complete,previous):
        best[certificate.mode]=certificate


def automatic_pair(args,coarse,record,folder,budget,best,checkpoint):
    """Return eligibility/result metadata; expected rejection is a fallback."""
    budget.check('automatic_candidate_geometry')
    outline=core.np.asarray(record['source_case']['outline'],dtype=float)
    if outline.shape!=(6,2):raise ValueError('A six-vertex slope is required')
    height=float(outline[4,1]-outline[3,1])
    upper_mesh,angle,change=wedge.wedge_mesh(record,.001*height)
    if upper_mesh.physical_model_id!=coarse.physical_model_id:
        raise ValueError('The automatic mesh changed the original physical model')
    change['height']=height
    change['depth_over_height']=.001
    critical=wedge.brentq(lambda fs:wedge.strength(upper_mesh,0,fs)-angle,
                         1e-4,1000.,xtol=1e-13)
    change['analytical_wedge_threshold_candidate_only']=critical
    change['analytical_threshold_is_a_published_bound']=False
    driver.save_mesh(folder/'candidate_upper_mesh.npz',upper_mesh)
    driver.write(folder/'mesh_change.json',change)
    budget.check('automatic_explicit_original_audit')
    field,audit=wedge.explicit_witness(upper_mesh,angle,critical+5e-5)
    core.np.savez_compressed(folder/'explicit_upper_candidate.npz',x=field,fs=audit['fs'],mode='upper',q=8)
    driver.write(folder/'explicit_upper_audit.json',audit)
    warm=None
    if audit.get('physical_pass') and audit.get('witness'):
        warm=driver.Certificate('upper',audit['fs'],upper_mesh,field,audit,
                                'fresh automatic explicit c0 wedge',False)
    upper,upper_root=driver.search(upper_mesh,'upper',(max(1e-5,critical-.002),audit['fs']),
        folder/'upper',args.case+'_automatic_upper',args.tolerance,args.solve_seconds,
        initial_witness=warm,budget=budget)
    retain(best,checked_certificate(upper,args.case))
    checkpoint('automatic_upper_complete')
    if upper is None or not upper.root_complete:
        return {'eligible':True,'complete':False,'reason':'Automatic upper own-root incomplete','mesh_change':change}
    budget.check('automatic_fresh_lower')
    lower_mesh=coarse
    lower_change={'method':'original_starting_mesh','elements':len(coarse.tri)}
    if args.fan_lower:
        from mesh_accuracy import angular_fan_refine
        candidate,lower_change=angular_fan_refine(coarse,record,scope='star')
        if len(candidate.tri)<=args.max_elements:lower_mesh=candidate
        else:lower_change['skipped']='element_cap'
    if lower_mesh.physical_model_id!=coarse.physical_model_id:
        raise ValueError('The fresh lower mesh changed the original physical model')
    driver.save_mesh(folder/'candidate_lower_mesh.npz',lower_mesh)
    driver.write(folder/'lower_mesh_change.json',lower_change)
    lower,lower_root=driver.search(lower_mesh,'lower',(.4,1.3),folder/'lower',
        args.case+'_automatic_lower',args.tolerance,args.solve_seconds,budget=budget)
    retain(best,checked_certificate(lower,args.case))
    checkpoint('automatic_lower_complete')
    return {'eligible':True,'complete':bool(lower and lower.root_complete and upper.root_complete),
            'upper_root':upper_root,'lower_root':lower_root,'mesh_change':change,'lower_mesh_change':lower_change}


def run(args):
    global _UPPER_DIAGNOSTICS
    if (not driver.finite(args.target_gap,args.case_seconds,args.solve_seconds,args.tolerance,args.theta) or
            not 0<args.target_gap<2 or args.case_seconds<=0 or args.solve_seconds<=0 or
            not 0<args.tolerance<=1e-4 or not 0<args.theta<=1 or args.cycles<0 or args.max_elements<=0):
        raise ValueError('Invalid bounded precision profile')
    _UPPER_DIAGNOSTICS=args.output.resolve()/args.case/'scaled_upper_attempts'
    driver.solve=scaled_upper_solve
    # Explicit seed/resume retains its honest scope and the original interface.
    if args.seed or args.resume or args.upper_only:
        hashes=implementation_hashes()
        report=driver.run(args)
        report['automatic_candidate']={'attempted':False,'reason':'Explicit saved-field or upper-only mode'}
        report['wrapper_source_sha256']=hashes
        if implementation_hashes()!=hashes:
            report.update(status='ERROR',target_met=False,reason='Wrapper source changed during execution')
        driver.write(args.output.resolve()/args.case/'summary.json',report)
        return report
    budget=driver.CaseBudget(args.case_seconds,args.solve_seconds)
    folder=args.output.resolve()/args.case
    if (folder/'summary.json').exists():raise ValueError('Use a fresh output directory or --resume')
    folder.mkdir(parents=True,exist_ok=True)
    trial=folder/'automatic_c0_wedge';trial.mkdir(exist_ok=True)
    hashes=implementation_hashes()
    driver.write(folder/'source_manifest.json',{'source_sha256':hashes})
    best={};automatic={'attempted':True};fallback=None

    def checkpoint(reason,finished=False):
        metadata={mode:certificate.publish(folder) for mode,certificate in best.items()}
        pair=set(best)=={'lower','upper'}
        width=upper_width=None
        if pair:
            lower,upper=best['lower'].fs,best['upper'].fs
            if not 0<lower<=upper:raise ValueError('Reversed or nonpositive endpoints')
            width=2*(upper-lower)/(upper+lower);upper_width=(upper-lower)/upper
        complete=bool(pair and all(c.root_complete for c in best.values()))
        budget_record=budget.record()
        unchanged=implementation_hashes()==hashes
        met=bool(pair and width<=args.target_gap and complete and not budget_record['exhausted'] and unchanged)
        status=('ERROR' if not unchanged else 'BUDGET_EXHAUSTED' if budget_record['exhausted'] else
                'PASS' if finished and met else 'PARTIAL' if finished else 'RUNNING')
        report={'case':args.case,'status':status,'reason':reason,**metadata,'target_met':met,
            'target_gap':args.target_gap,'root_complete':complete,'relative_width_symmetric':width,
            'relative_width_upper':upper_width,'two_percent_width_pass':bool(pair and width<=.02),
            'one_percent_width_pass':bool(pair and width<=.01),'case_budget':budget_record,
            'budget_expired':budget_record['exhausted'],'source_unchanged':unchanged,
            'automatic_candidate':automatic,'fallback':fallback,
            'execution_scope':{'kind':'fresh_original_model_with_exact_upper_velocity_units_v3',
                'description':'Fresh original coordinates and loads; exact LP variable substitution and original q8 audits; no historical seed or Fs',
                'q':8,'root_fs_tolerance':args.tolerance,'not_an_excel_timing_test':True},
            'floating_tolerance_certificate':True,'strict_interval_arithmetic_proof':False,
            'environment':{'python':platform.python_version(),'numpy':core.np.__version__,
                'clarabel':core.clarabel.__version__,'threads':1,'source_sha256':hashes},
            'profile':{'case_seconds':args.case_seconds,'solve_seconds':args.solve_seconds,
                'target_gap':args.target_gap,'cycles':args.cycles,'max_elements':args.max_elements,
                'theta':args.theta,'fan_lower':args.fan_lower,'fs_tolerance':args.tolerance}}
        driver.write(folder/'case_budget.json',budget_record)
        driver.write(folder/'summary.json',report)
        return report

    checkpoint('starting_fresh_candidate')
    record=None
    try:
        budget.check('starting_model')
        coarse,record=driver.load_model(args.case)
        try:
            automatic.update(automatic_pair(args,coarse,record,trial,budget,best,checkpoint))
        except driver.BudgetExceeded:raise
        except (ValueError,RuntimeError,AssertionError,KeyError,IndexError) as error:
            automatic.update(eligible=False if not (trial/'mesh_change.json').exists() else True,
                             complete=False,reason=type(error).__name__+': '+str(error))
        driver.write(trial/'automatic_result.json',automatic)
        report=checkpoint('automatic_candidate_complete')
        if report['target_met']:return checkpoint('target_met_with_fresh_automatic_pair',True)
        budget.check('generic_fallback')
        remaining=budget.remaining()
        fallback_args=copy.copy(args)
        fallback_args.output=folder/'generic_fallback'
        fallback_args.case_seconds=remaining
        fallback={'scope':'fresh generic restart; no historical or automatic witness seed',
                  'allocated_case_seconds':remaining,'output':str(fallback_args.output/args.case)}
        checkpoint('starting_generic_fallback')
        try:
            fallback_report=driver.run(fallback_args)
            fallback.update(status=fallback_report.get('status'),case_budget=fallback_report.get('case_budget'),
                            target_met=fallback_report.get('target_met',False))
        except Exception as error:
            fallback.update(status='ERROR',error=type(error).__name__+': '+str(error))
        for mode in ('lower','upper'):
            try:retain(best,read_certificate(fallback_args.output/args.case,mode,record,args.case))
            except (ValueError,OSError,KeyError,TypeError,IndexError) as error:
                fallback.setdefault('rejected_candidates',[]).append({'mode':mode,'reason':str(error)})
        return checkpoint('fresh_generic_fallback_finished',True)
    except driver.BudgetExceeded as error:
        automatic.setdefault('reason',str(error))
        driver.write(trial/'automatic_result.json',automatic)
        return checkpoint(str(error),True)
    except Exception as error:
        report=checkpoint(type(error).__name__+': '+str(error),True)
        report.update(status='ERROR',target_met=False)
        driver.write(folder/'summary.json',report)
        raise


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--case',required=True)
    parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--seed',type=Path)
    parser.add_argument('--resume',action='store_true')
    parser.add_argument('--recheck',action='store_true')
    parser.add_argument('--fan-lower',action='store_true')
    parser.add_argument('--upper-only',action='store_true')
    parser.add_argument('--cycles',type=int,default=8)
    parser.add_argument('--max-elements',type=int,default=32000)
    parser.add_argument('--theta',type=float,default=.6)
    parser.add_argument('--tolerance',type=float,default=1e-4)
    parser.add_argument('--target-gap',type=float,default=.02)
    parser.add_argument('--case-seconds',type=float,default=900.)
    parser.add_argument('--solve-seconds',type=float,default=30.)
    run(parser.parse_args())


if __name__=='__main__':main()
