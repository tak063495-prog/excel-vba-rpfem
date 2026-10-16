"""Budgeted Python26 policy; no certificate is inferred from a solver status.

Run with Python -X utf8. Bounds carry their own mesh and field. Failed trials
never overwrite a witness, and numerical signs only bracket the discrete root.
"""
from __future__ import annotations
import argparse
from dataclasses import dataclass
from pathlib import Path
import json
import hashlib
import platform
import struct
import time
import experiment_core as core
from experiment_core import np, Prepared, LowerRefresh, evaluate, solve, write as _core_write

REPO = Path(__file__).resolve().parents[3]
CATALOG = REPO / 'docs/benchmarks/G1R17_native26_20261013/data/inputs'
GEOMETRY = REPO / 'versions/G1R17/results/geometry25'
ACCURATE = {'equilibrate_enable': False}
SOURCE_PATHS=[Path(__file__),Path(core.__file__),Path(__file__).with_name('reduction.py'),
              Path(__file__).with_name('zero_phi_transport.py'),Path(__file__).with_name('mesh_accuracy.py'),
              Path(__file__).with_name('original_farkas.py'),
              Path(__file__).with_name('original_dual.py'),
              Path(__file__).with_name('benchmark_models.py'),
              Path(__file__).parent/'vendor/independent_audit.py',
              Path(__file__).parent/'vendor/model_support/lower_assembly.py',
              Path(__file__).parent/'vendor/model_support/upper_assembly.py']


def source_hashes():
    """Freeze the active copied implementation, never an old pipeline's paths."""
    paths=sorted(set(SOURCE_PATHS)|set(Path(__file__).parent.rglob('*.py')))
    return {p.relative_to(REPO).as_posix():hashlib.sha256(p.read_bytes()).hexdigest()
            for p in paths if p.is_file() and '__pycache__' not in p.parts}


SOURCE_HASHES=source_hashes()


def write(path, obj):
    """A watchdog must not expose a half-written JSON checkpoint."""
    path=Path(path)
    temporary=path.with_name(path.name+'.partial')
    _core_write(temporary,obj)
    temporary.replace(path)


class BudgetExceeded(RuntimeError):
    pass


class CaseBudget:
    """Cooperative whole-case wall budget, shared by every root and recovery.

    C/Fortran assembly, restoration and LSMR cannot be interrupted here. The
    orchestrator must also supervise the process; time_limit alone is not a
    whole-case wall deadline. An expired call may finish publishing evidence,
    but cannot start another solve or produce a PASS result.
    """
    def __init__(self, case_seconds=900., solve_seconds=30., clock=None):
        if (not finite(case_seconds,solve_seconds) or
                case_seconds<=0 or solve_seconds<=0):
            raise ValueError('Case and solve budgets must be finite and positive')
        self.clock=clock or time.perf_counter
        self.case_seconds=float(case_seconds)
        self.solve_seconds=float(solve_seconds)
        self.started=self.clock()
        self.deadline=self.started+self.case_seconds
        self.allocations=[]
        self.expired_at=None

    def remaining(self):
        return max(0.,self.deadline-self.clock())

    def check(self, stage):
        if self.remaining()<=0:
            self.expired_at=self.expired_at or stage
            raise BudgetExceeded('Whole-case wall budget exhausted before '+stage)

    def allowance(self, stage, maximum=None):
        self.check(stage)
        limit=min(self.solve_seconds,self.remaining(),
                  float(maximum) if maximum is not None else self.solve_seconds)
        self.allocations.append({'stage':stage,'time_limit':limit,
                                 'elapsed_seconds':self.clock()-self.started})
        return limit

    def record(self):
        return {'case_seconds':self.case_seconds,'solve_seconds':self.solve_seconds,
                'elapsed_seconds':self.clock()-self.started,
                'remaining_seconds':self.remaining(),'exhausted':self.remaining()<=0,
                'expired_at':self.expired_at,'solver_time_allocations':self.allocations,
                'scope':'cooperative wall deadline; OS watchdog required for in-flight native work'}


def load_model(case):
    # Install once in the copied core so all adaptive and saved-mesh rebuilds
    # share the same dry/coupled model builder, including D08.
    from benchmark_models import install_model_builder, load_model as load_benchmark
    install_model_builder()
    return load_benchmark(case)


def load_saved_mesh(path, rec):
    from benchmark_models import install_model_builder
    install_model_builder()
    with np.load(path) as data:
        return core.mesh_from_arrays(data['xy'].copy(), data['tri'].copy(), data['material_id'].copy(), rec)


def save_mesh(path, mesh):
    path=Path(path)
    temporary=path.with_name(path.stem+'.partial.npz')
    np.savez_compressed(temporary, xy=mesh.xy, tri=mesh.tri, material_id=mesh.material_id)
    temporary.replace(path)


@dataclass
class Certificate:
    mode: str
    fs: float
    mesh: object
    field: object
    result: dict
    source: str
    root_complete: bool = False

    def publish(self, folder):
        stem = folder / f'best_{self.mode}'
        save_mesh(stem.with_name(stem.name + '_mesh.npz'), self.mesh)
        temporary=stem.with_name(stem.name+'_field.partial.npz')
        np.savez_compressed(temporary,x=self.field,fs=self.fs,mode=self.mode,q=8)
        temporary.replace(stem.with_name(stem.name+'_field.npz'))
        metadata={'case':self.mesh.provenance['case'],'mode':self.mode,'q':8,'fs': self.fs, 'mesh_id': self.mesh.mesh_id, 'physical_model_id': self.mesh.physical_model_id,
                'elements': len(self.mesh.tri), 'source': self.source, 'root_complete': self.root_complete,
                'mesh_path': stem.name + '_mesh.npz', 'field_path': stem.name + '_field.npz', 'audit': self.result}
        write(folder/f'best_{self.mode}_certificate.json',metadata)
        return metadata


def finite(*values):
    try:
        return bool(np.isfinite(np.asarray(values,dtype=float)).all())
    except (TypeError,ValueError):
        return False


def sign(result, mode):
    target = core.LOWER_TARGET if mode == 'lower' else core.UPPER_TARGET
    if result.get('witness',False) and result.get('physical_pass',False) and finite(result.get('value')):
        return 1 if mode == 'lower' else -1
    if result.get('numerical_gate',False) and finite(result.get('value')):
        return 1 if result['value'] >= target else -1
    if not finite(result.get('dual_stationarity'),result.get('dual_cone_error'),result.get('dual_objective')):
        return None
    if result.get('dual_stationarity', 1) > 1e-7 or result.get('dual_cone_error', 1) > 1e-8:
        return None
    if mode == 'lower' and -result['dual_objective'] <= target - 2e-6:
        return -1
    if mode == 'upper' and result['dual_objective'] >= target + 2e-6:
        return 1
    return None


def physical_attempt_rank(result, mode):
    """Prefer the better admissible objective only after all existing gates.

    A low-quality first admissible iterate must not hide a later iterate that
    permits fixed-load recovery. Values from rejected fields have no rank.
    This selection never promotes a field or supplies a numerical root sign.
    """
    if mode not in ('lower','upper'):
        raise ValueError('Lower or upper mode required')
    admissible=bool(result.get('physical_pass') and finite(result.get('value')))
    quality=float(result['value'])*(1 if mode=='lower' else -1) if admissible else 0.
    return (bool(result.get('witness')),bool(result.get('physical_pass')),
            sign(result,mode) is not None,bool(result.get('numerical_gate')),quality)


def audit_initial_dual_guide(mesh, mode, guide):
    """Reassemble and re-audit an original vector; never create a primal field."""
    required = {'case': mesh.provenance['case'], 'mode': mode, 'q': 8,
                'mesh_id': mesh.mesh_id, 'physical_model_id': mesh.physical_model_id,
                'policy': mesh.provenance['psi_policy']}
    if any(guide.get(key) != value for key, value in required.items()):
        raise ValueError('An initial original dual guide has different model or mesh identity')
    fs = guide.get('fs')
    if not finite(fs) or not isinstance(fs, (int, float, np.integer, np.floating)) or fs <= 0:
        raise ValueError('An initial original dual guide requires a finite positive Fs')
    prepared = Prepared(mesh, float(fs), mode, 8, reduction='none')
    from original_dual import audit_original_dual
    dual = np.asarray(guide.get('dual'), dtype=float).copy()
    evidence = audit_original_dual(prepared, dual)
    if not evidence['accepted']:
        raise ValueError('The unchanged original dual gate rejected the initial guide: '+evidence.get('reason', ''))
    result = dict(evidence, case=required['case'], mode=mode, q=8, fs=float(fs),
        mesh_id=mesh.mesh_id, physical_model_id=mesh.physical_model_id, policy=required['policy'],
        physical_pass=False, witness=False, numerical_gate=False, value=0.,
        dual_guide_only=True, origin='independently_reaudited_initial_original_dual',
        scope='discrete_root_search_only; no primal field or continuum bound')
    return float(fs), result, dual


def search(mesh, mode, initial, folder, label, tolerance=1e-4, seconds=120, initial_witness=None,
           initial_dual_guides=None, budget=None, bisection=False):
    """Root audit with an optional shared case budget; old callers stay valid."""
    if not finite(seconds,tolerance) or seconds<=0 or not 0<tolerance<=1e-4:
        raise ValueError('Positive solve seconds and unchanged Fs tolerance <=1e-4 required')
    folder.mkdir(parents=True, exist_ok=True)
    started = time.perf_counter()
    def check(stage):
        if budget is not None:budget.check(label+':'+stage)
    def allowance(stage, maximum=seconds):
        return budget.allowance(label+':'+stage,maximum) if budget is not None else maximum
    def can_continue():
        return budget is None or budget.remaining()>0
    check('initial_prepare')
    refresh = LowerRefresh(mesh, 'affine') if mode == 'lower' else None
    if mode=='upper' and np.all(mesh.materials[:,1:3]==0) and not np.any(mesh.body[:,:,0]):
        from zero_phi_transport import PhiZeroUpperRefresh
        refresh=PhiZeroUpperRefresh(mesh,fs=1.,q=8,reduction='affine')
    target = core.LOWER_TARGET if mode == 'lower' else core.UPPER_TARGET
    cache, trials = {}, []
    best = None
    if initial_witness is not None:
        if initial_witness.mode!=mode or initial_witness.mesh.mesh_id!=mesh.mesh_id or not initial_witness.result['witness']:
            raise ValueError('An initial witness must own this exact mesh and side')
        result=dict(initial_witness.result)
        result.pop('discrete_root_search',None)
        result['origin']='independently_reaudited_same_mesh_witness'
        cache[initial_witness.fs]=(result,initial_witness.field.copy())
        trials.append(result)
        best=Certificate(mode,initial_witness.fs,mesh,initial_witness.field.copy(),result,label)
    for number, guide in enumerate(initial_dual_guides or []):
        check('initial_original_dual_audit')
        fs, result, dual = audit_initial_dual_guide(mesh, mode, guide)
        if fs in cache:
            raise ValueError('An initial dual guide conflicts with a stored physical witness')
        proof_path = folder/(label+f'_original_dual_guide_{number+1}.npz')
        np.savez_compressed(proof_path, dual=dual, fs=fs, mode=mode, q=8,
            case=mesh.provenance['case'], mesh_id=mesh.mesh_id,
            physical_model_id=mesh.physical_model_id, policy=mesh.provenance['psi_policy'])
        result.update(dual_proof_path=proof_path.name,
                      dual_proof_sha256=hashlib.sha256(proof_path.read_bytes()).hexdigest())
        cache[fs] = result, None
        trials.append(result)
    if initial_dual_guides:
        write(folder/(label+'_trials.json'), trials)

    def point(fs):
        nonlocal best
        fs = float(fs)
        if fs in cache:
            return cache[fs]
        check('point_prepare')
        prepared = refresh.at(fs) if refresh else Prepared(mesh, fs, mode, 8, reduction='affine')
        attempts = []
        dual_evidence = []
        chosen = None
        # Choose physical evidence before numerical optimality. A later failed
        # recovery cannot replace an earlier, independently audited witness.
        for options,scaled in [(ACCURATE,False), (dict(ACCURATE, iterative_refinement_max_iter=30,
                                      iterative_refinement_abstol=1e-15, iterative_refinement_reltol=1e-15),
                                      False),
                        (dict(ACCURATE, tol_gap_abs=1e-13, tol_gap_rel=1e-13, tol_feas=1e-13,
                              iterative_refinement_max_iter=30,iterative_refinement_abstol=1e-15,
                              iterative_refinement_reltol=1e-15),True),
                        ({'equilibrate_enable': True},False)]:
            if not can_continue():break
            allocated=allowance('LP_attempt_'+str(len(attempts)+1))
            result, field, _ = solve(prepared, dict(options, time_limit=allocated), max_iter=350,cone_scaling=scaled,primal_polish=True)
            attempts.append({k: result[k] for k in ['status', 'physical_pass', 'witness', 'numerical_gate', 'end_to_end_seconds']})
            dual_evidence.append({k: result.get(k) for k in ['dual_objective','dual_stationarity','dual_cone_error',
                                    'reduced_dual_objective','reduced_dual_stationarity','reduced_dual_cone_error']})
            rank = physical_attempt_rank(result,mode)
            if chosen is None or rank > chosen[0]:
                chosen = rank, result, field
            if sign(result,mode) is not None:
                break
        if chosen is None:
            check('LP_attempt')
            raise RuntimeError('No solve attempt produced a result')
        _, result, field = chosen
        # Separate iterates at the same Fs may provide the admissible primal
        # and the admissible dual. Their gap need not be small for a dual to
        # screen a discrete root; the published bound is still only the primal.
        if sign(result,mode) is None:
            for evidence in dual_evidence:
                merged=dict(result,**evidence)
                if sign(merged,mode) is not None:
                    result=merged
                    result['dual_from_recovery_at_same_fs']=True
                    break
        original=getattr(prepared,'original',None)
        if can_continue() and mode=='upper' and sign(result,mode) is None and original is not None and np.all(original[-1]==0):
            from original_farkas import solve_original_farkas
            evidence=solve_original_farkas(prepared,seconds=allowance('original_farkas',min(seconds,30.)))
            result['original_farkas_recovery']=evidence
            if evidence['accepted']:
                result.update({key:evidence[key] for key in ['dual_objective','dual_stationarity','dual_cone_error']})
                result['numerical_gate']=False
                result['dual_from_original_farkas_at_same_fs']=True
        if can_continue() and sign(result,mode) is None and original is not None:
            from original_dual import solve_original_dual
            evidence=solve_original_dual(prepared,seconds=allowance('original_dual',min(seconds,30.)))
            result['original_dual_feasibility_recovery']=evidence
            if evidence['accepted']:
                result.update({key:evidence[key] for key in ['dual_objective','dual_stationarity','dual_cone_error']})
                result['numerical_gate']=False
                result['dual_from_original_dual_feasibility_at_same_fs']=True
        if can_continue() and mode=='lower' and not result.get('witness') and sign(result,mode)!=-1 and result.get('value',0)>=.98:
            # A feasibility solve avoids maximizing a nearly singular load
            # variable. Its failure is UNKNOWN; success is only a primal bound.
            feasibility=Prepared(mesh,fs,'lower',8,reduction='affine',fixed_load=1.000003)
            options=dict(ACCURATE,iterative_refinement_max_iter=30,iterative_refinement_abstol=1e-15,
                         iterative_refinement_reltol=1e-15,time_limit=allowance('fixed_load_feasibility'))
            witness,wx,_=solve(feasibility,options,max_iter=350,cone_scaling=True,primal_polish=True)
            attempts.append({'fixed_load':1.000003,**{k:witness[k] for k in ['status','physical_pass','witness','end_to_end_seconds']}})
            if witness['witness']:
                # The fixed-load dual is for a different objective and must
                # never screen the load-maximization root.
                witness['numerical_gate']=False
                witness['reduced_numerical_gate']=False
                for key in list(witness):
                    if 'dual' in key:del witness[key]
                witness['fixed_load_feasibility_only']=True
                result,field=witness,wx
        result.update(attempts=attempts, original_dual_recovery_attempts=dual_evidence,
                      search_sign=sign(result, mode), fs=fs)
        cache[fs] = result, field
        trials.append(result)
        if result['witness'] and (best is None or (fs > best.fs if mode == 'lower' else fs < best.fs)):
            best = Certificate(mode, fs, mesh, field.copy(), result, label)
            best.publish(folder)
        write(folder / (label + '_trials.json'), trials)
        print('POINT', label, fs, result['value'], result['status'], result['physical_pass'], result['witness'], flush=True)
        return result, field

    lo, hi = map(float, initial)
    if not finite(lo,hi) or not 0<lo<hi:
        raise ValueError('An initial bracket must have finite positive, ordered endpoints')
    lr, hr = {}, {}
    complete = False
    budget_error = None
    try:
        if initial_witness is not None and lo<initial_witness.fs<hi:
            if mode=='lower':lo=initial_witness.fs
            else:hi=initial_witness.fs
        lr, _ = point(lo)
        hr, _ = point(hi)
        if sign(lr,mode) is None or sign(hr,mode) is None:
            # Distant initial strengths can fail numerically even on a nested
            # mesh. Probe near the preceding endpoint before abandoning it.
            middle=(lo+hi)/2
            for fs in [middle,middle-2e-5,middle+2e-5]:
                result,_=point(fs)
                if sign(result,mode)==1 and (sign(lr,mode)!=1 or fs>lo):
                    lo,lr=fs,result
                elif sign(result,mode)==-1 and (sign(hr,mode)!=-1 or fs<hi):
                    hi,hr=fs,result
                if sign(lr,mode)==1 and sign(hr,mode)==-1 and lo<hi:
                    break
        # Expand just the missing side, keeping every physical witness obtained.
        for _ in range(5):
            if sign(lr, mode) == 1 and sign(hr, mode) == -1:
                break
            span = hi - lo
            if sign(lr, mode) == -1:
                next_lo=max(1e-5,lo-span)
                next_lr,_=point(next_lo)
                lo,lr=next_lo,next_lr
            elif sign(hr, mode) == 1:
                next_hi=hi+span
                next_hr,_=point(next_hi)
                hi,hr=next_hi,next_hr
            else:
                break
        complete = False
        if initial_witness is not None and lo<initial_witness.fs<hi:
            warm_result,_=cache[initial_witness.fs]
            if mode=='lower':lo,lr=initial_witness.fs,warm_result
            else:hi,hr=initial_witness.fs,warm_result
        if sign(lr, mode) == 1 and sign(hr, mode) == -1 and lo<hi:
            for _ in range(32):
                if hi - lo <= tolerance * max(1., (hi + lo) / 2):
                    complete = True
                    break
                denominator = 0. if bisection or hr.get('dual_guide_only') or lr.get('dual_guide_only') else hr['value'] - lr['value']
                fs = (lo + hi) / 2 if denominator == 0 or not finite(denominator,lr['value'],hr['value']) else lo + (target - lr['value']) * (hi - lo) / denominator
                fs = float(np.clip(fs, lo + .2 * (hi - lo), hi - .2 * (hi - lo)))
                result, _ = point(fs)
                if sign(result, mode) is None:
                    # Isolated numerical failures near a face transition need not
                    # erase a signed bracket. Probe bounded interior alternatives;
                    # each sign still requires the unchanged original audit.
                    for alternative in [(lo+hi)/2,lo+.25*(hi-lo),lo+.75*(hi-lo)]:
                        result,_=point(alternative)
                        fs=alternative
                        if sign(result,mode) is not None:
                            break
                s = sign(result, mode)
                if s is None:
                    break
                if s == 1:
                    lo, lr = fs, result
                else:
                    hi, hr = fs, result
    except BudgetExceeded as exc:
        budget_error=str(exc)
    if budget is not None and budget.remaining()<=0:
        complete=False
        budget_error=budget_error or 'Whole-case budget exhausted during an in-flight call'
    endpoint = lo if mode == 'lower' else hi
    if best is not None:
        best.root_complete = bool(complete and best.fs == endpoint)
    summary = {'mode': mode, 'root_complete': bool(best and best.root_complete), 'bracket': [lo, hi],
               'bracket_width': hi - lo, 'fs_tolerance': tolerance, 'endpoint': best.fs if best else None,
               'trials': len(trials), 'seconds': time.perf_counter() - started,
               'mesh_id':mesh.mesh_id,'physical_model_id':mesh.physical_model_id,'elements':len(mesh.tri),
               'q':8,'case':mesh.provenance['case'],'policy':mesh.provenance['psi_policy'],
               'bracket_signs':[sign(lr,mode),sign(hr,mode)],
               'opposite_numerical_endpoint_is_continuum_bound': False,
               'budget_exhausted':budget_error is not None,'budget_error':budget_error,
               'case_budget':budget.record() if budget is not None else None,
               'proposal_method':'bounded_bisection' if bisection else 'bounded_secant'}
    write(folder / (label + '_summary.json'), summary)
    if best is not None:
        best.result['discrete_root_search'] = summary
        best.publish(folder)
    return best, summary


def run(args):
    """Execute the isolated two-percent policy and always publish a case row."""
    args.target_gap=getattr(args,'target_gap',.02)
    args.case_seconds=getattr(args,'case_seconds',900.)
    args.independent_lower=getattr(args,'independent_lower',False)
    args.lower_theta=getattr(args,'lower_theta',args.theta)
    args.upper_indicator=getattr(args,'upper_indicator','mixed')
    args.kinematic_fraction=getattr(args,'kinematic_fraction',.2)
    args.initial_span=getattr(args,'initial_span',.005)
    args.initial_mesh=getattr(args,'initial_mesh',None)
    args.root_method=getattr(args,'root_method','bisection')
    if (not finite(args.tolerance,args.theta,args.solve_seconds,args.case_seconds,args.target_gap) or
            not 0<args.tolerance<=1e-4 or not 0<args.theta<=1 or
            not 0<args.target_gap<2 or args.solve_seconds<=0 or args.case_seconds<=0 or
            args.cycles<0 or args.max_elements<=0):
        raise ValueError('Finite positive budgets, target gap in (0,2), and Fs tolerance <=1e-4 required')
    if (not finite(args.lower_theta,args.kinematic_fraction,args.initial_span) or
            not 0<args.lower_theta<=1 or not 0<=args.kinematic_fraction<=1 or args.initial_span<=0 or
            args.upper_indicator not in ('mixed','kinematic') or args.root_method not in ('bisection','secant')):
        raise ValueError('Finite positive refinement/span parameters and a convex indicator fraction required')
    from adaptive_policy import stable_prolong, upper_refinement_score
    budget=CaseBudget(args.case_seconds,args.solve_seconds)
    folder=args.output.resolve()/args.case
    folder.mkdir(parents=True,exist_ok=True)
    previous_summary=None
    if (folder/'summary.json').exists():
        if not args.resume:
            raise ValueError('Existing case output requires --resume or a fresh output directory')
        previous_summary=json.loads((folder/'summary.json').read_text(encoding='utf8'))
        args.seed=folder
    frozen_sources=source_hashes()
    initial_mesh_sha=hashlib.sha256(args.initial_mesh.read_bytes()).hexdigest() if args.initial_mesh else None
    write(folder/'source_manifest.json',{'source_sha256':frozen_sources})
    best,history={},[]
    scope={'kind':'saved_witness_reaudit_and_new_python_solves' if args.seed or args.resume else 'cold_start_new_python_policy',
           'description':'New Python solver policy on frozen benchmark starting geometry; no G1R19 Excel cold-start timing claim',
           'saved_witness_seed_requested':bool(args.seed or args.resume),
           'mesh_only_input':str(args.initial_mesh) if args.initial_mesh else None,
           'mesh_only_input_sha256':initial_mesh_sha,
           'quadrature':8,'root_fs_tolerance':args.tolerance,'target_relative_width_symmetric':args.target_gap}
    mesh=None
    stop_reason='cycle_budget_reached'

    def checkpoint(status='RUNNING',reason=None):
        metadata={mode:cert.publish(folder) for mode,cert in best.items()}
        pair=set(best)=={'lower','upper'}
        width=upper_width=None
        if pair:
            lb,ub=best['lower'].fs,best['upper'].fs
            if not finite(lb,ub) or not 0<lb<=ub:
                raise ValueError('Reversed, nonfinite or nonpositive bounds')
            if best['lower'].mesh.physical_model_id!=best['upper'].mesh.physical_model_id:
                raise ValueError('Independent bounds refer to different physical models')
            width=2*(ub-lb)/(lb+ub)
            upper_width=(ub-lb)/ub
        roots=bool(pair and all(c.root_complete for c in best.values()))
        budget_state=budget.record()
        current_sources=source_hashes()
        unchanged=current_sources==frozen_sources
        if args.initial_mesh:
            unchanged=bool(unchanged and hashlib.sha256(args.initial_mesh.read_bytes()).hexdigest()==initial_mesh_sha)
        met=bool(pair and width<=args.target_gap and roots and not budget_state['exhausted'] and unchanged)
        if not unchanged:
            status='ERROR';reason='Implementation source changed during the case'
        elif budget_state['exhausted']:
            status='BUDGET_EXHAUSTED';reason=reason or 'Whole-case wall budget exhausted'
        elif status=='FINISHED':
            status='PASS' if met else 'PARTIAL'
        report={'case':args.case,'status':status,'reason':reason,'execution_scope':scope,
                **metadata,'relative_width_upper':upper_width,'relative_width_symmetric':width,
                'target_gap':args.target_gap,'two_percent_width_pass':bool(pair and width<=.02),
                'one_percent_width_pass':bool(pair and width<=.01),'root_complete':roots,'target_met':met,
                'floating_tolerance_certificate':True,'strict_interval_arithmetic_proof':False,
                'case_budget':budget_state,'source_unchanged':unchanged,
                'environment':{'python':platform.python_version(),'platform':platform.platform(),
                    'numpy':np.__version__,'clarabel':core.clarabel.__version__,'threads':1,
                    'source_sha256':frozen_sources,'source_sha256_at_checkpoint':current_sources},
                'profile':{'cycles':args.cycles,'max_elements':args.max_elements,'theta':args.theta,
                    'fan_lower':args.fan_lower,'solve_seconds':args.solve_seconds,
                    'case_seconds':args.case_seconds,'target_gap':args.target_gap,
                    'fs_tolerance':args.tolerance},'history':history}
        report['profile'].update(independent_lower=args.independent_lower,lower_theta=args.lower_theta,
            upper_indicator=args.upper_indicator,kinematic_fraction=args.kinematic_fraction,
            initial_span=args.initial_span,precision_local_prolongation=True,stop_after_each_mode=True,
            root_method=args.root_method)
        write(folder/'case_budget.json',budget_state)
        write(folder/'summary.json',report)
        return report

    checkpoint()
    try:
        budget.check('initial_model')
        mesh,rec=load_model(args.case)
        if args.initial_mesh:
            if args.seed or args.resume:
                raise ValueError('A mesh-only cold start and a witness seed are mutually exclusive')
            if args.initial_mesh.suffix.lower()=='.npz':
                mesh=load_saved_mesh(args.initial_mesh,rec)
            else:
                raw=args.initial_mesh.read_bytes()
                if len(raw)<8:raise ValueError('Invalid mesh-only input header')
                nn,ne=struct.unpack_from('<ii',raw)
                if not 3<=nn<=100000 or not 1<=ne<=220000 or len(raw)!=8+16*nn+16*ne:
                    raise ValueError('Invalid finite mesh-only input extent')
                xy=np.frombuffer(raw,dtype='<f8',count=2*nn,offset=8).reshape(nn,2).copy()
                cells=np.frombuffer(raw,dtype='<i4',count=4*ne,offset=8+16*nn).reshape(ne,4).copy()
                if not np.isfinite(xy).all() or np.any(cells[:,:3]<0) or np.any(cells[:,:3]>=nn):
                    raise ValueError('Invalid mesh-only coordinates or connectivity')
                mesh=core.mesh_from_arrays(xy,cells[:,:3],cells[:,3],rec)
            if len(mesh.tri)>args.max_elements:raise ValueError('Initial mesh exceeds the element cap')
        budget.check('initial_model_complete')
        velocity=None
        if getattr(args,'upper_only',False) and not args.seed:
            raise ValueError('Upper-only refinement requires independently audited lower and upper seeds')
        seed_meta = {}
        if args.seed:
            if args.seed.resolve()==folder and previous_summary is not None:
                seed_meta=previous_summary
                history=seed_meta.get('history',[])
            elif (args.seed/'summary.json').exists():
                seed_meta=json.loads((args.seed/'summary.json').read_text(encoding='utf8'))
                history=seed_meta.get('history', [])
            elif (args.seed/'import_audit.json').exists():
                seed_meta=json.loads((args.seed/'import_audit.json').read_text(encoding='utf8'))['certificates']
            mesh = load_saved_mesh(args.seed / 'best_upper_mesh.npz', rec)
            for mode in ['lower', 'upper']:
                budget.check('seed_'+mode+'_original_audit')
                m = load_saved_mesh(args.seed / f'best_{mode}_mesh.npz', rec)
                with np.load(args.seed / f'best_{mode}_field.npz') as d:
                    x, fs = d['x'].copy(), float(d['fs'])
                    if str(d['mode'])!=mode or int(d['q'])!=8:
                        raise ValueError('Seed mode or quadrature differs from this run')
                result = evaluate(Prepared(m, fs, mode, 8, reduction='none'), x)
                if not result['witness']:
                    raise ValueError(f'{mode} seed failed independent physical audit')
                metadata=seed_meta.get(mode,{})
                original_result=metadata.get('audit', {})
                root=original_result.get('discrete_root_search', {})
                source_path=Path(metadata.get('source',''))
                if not root and (source_path/'import_audit.json').is_file():
                    imported=json.loads((source_path/'import_audit.json').read_text(encoding='utf8'))['certificates'][mode]
                    if imported['fs']==fs and imported['mesh_id']==m.mesh_id:
                        metadata=imported
                        root=metadata['audit'].get('discrete_root_search',{})
                bracket=root.get('bracket',[])
                ordered=bool(len(bracket)==2 and finite(*bracket) and 0<bracket[0]<bracket[1])
                narrow=bool(ordered and bracket[1]-bracket[0]<=args.tolerance*max(1.,sum(bracket)/2))
                identity=bool(metadata.get('fs')==fs and metadata.get('mesh_id')==m.mesh_id and
                    metadata.get('physical_model_id')==m.physical_model_id and metadata.get('case')==args.case and
                    metadata.get('mode')==mode and metadata.get('q')==8 and root.get('mesh_id')==m.mesh_id and
                    root.get('physical_model_id')==m.physical_model_id and root.get('case')==args.case and
                    root.get('policy')==m.provenance['psi_policy'] and root.get('q')==8 and root.get('mode')==mode)
                complete=bool(identity and narrow and root.get('bracket_signs')==[1,-1] and
                    metadata.get('root_complete') and root.get('root_complete') and root.get('endpoint')==fs and
                    fs==(bracket[0] if mode=='lower' else bracket[1]))
                if root:
                    result['discrete_root_search']=root
                best[mode] = Certificate(mode, fs, m, x, result, str(args.seed), complete)
            velocity = best['upper'].field
        initial=(.4,1.3)
        offset=max((r['cycle'] for r in history),default=0)+int(args.recheck)
        checkpoint()
        for step in range(args.cycles+1):
            budget.check('adaptive_cycle_'+str(step))
            cycle=offset+step
            if step or velocity is None or args.recheck:
                if velocity is not None and (step or not args.recheck):
                    parent=mesh
                    if args.upper_indicator=='mixed':
                        fs=best['upper'].fs if 'upper' in best else 1.
                        score,indicator_info=upper_refinement_score(mesh,velocity,fs,args.kinematic_fraction)
                    else:
                        score=core.indicator(mesh,velocity);indicator_info={'method':'existing_kinematic'}
                    mesh,refinement=core.local_refine(mesh,rec,score,args.theta)
                    refinement['indicator']=indicator_info
                    budget.check('upper_prolongation')
                    velocity=stable_prolong(parent,mesh,velocity,'upper')
                    if len(mesh.tri)>args.max_elements:
                        print('ELEMENT_CAP',len(mesh.tri),flush=True)
                        stop_reason='element_cap';break
                else:
                    refinement=None
                label=f'{args.case}_C{cycle}'
                save_mesh(folder/(label+'_mesh.npz'),mesh)
                row={'cycle':cycle,'elements':len(mesh.tri),'refinement':refinement}
                history.append(row)
                modes=['lower','upper'] if not getattr(args,'upper_only',False) or (args.recheck and step==0) else ['upper']
                for mode in modes:
                    budget.check(label+':'+mode)
                    solve_mesh=best[mode].mesh if args.recheck and step==0 and mode in best else mesh
                    lower_parent=None
                    if args.independent_lower and mode=='lower' and mode in best and not (args.recheck and step==0):
                        from mesh_accuracy import lower_stress_score
                        lower_parent=best[mode].mesh
                        score=lower_stress_score(lower_parent,best[mode].field,best[mode].fs)
                        solve_mesh,lower_info=core.local_refine(lower_parent,rec,score,args.lower_theta)
                        row['lower_independent_refinement']=lower_info
                    if args.fan_lower and mode=='lower' and not (args.recheck and step==0):
                        from mesh_accuracy import angular_fan_refine
                        fan_parent=solve_mesh
                        solve_mesh,fan_info=angular_fan_refine(fan_parent,rec,scope='star')
                        if lower_parent is not None:
                            ancestry=fan_parent.provenance['parent_element_ids']
                            solve_mesh.provenance['parent_element_ids']=[ancestry[e] for e in solve_mesh.provenance['parent_element_ids']]
                        if len(solve_mesh.tri)>args.max_elements:
                            solve_mesh=fan_parent;fan_info['skipped']='element_cap'
                        row['lower_angular_fan']=fan_info
                        save_mesh(folder/(label+'_lower_mesh.npz'),solve_mesh)
                    if len(solve_mesh.tri)>args.max_elements:
                        row[mode]={'skipped':'independent element cap','elements':len(solve_mesh.tri)}
                        continue
                    if mode in best:
                        old=best[mode].fs
                        delta=max(1e-4,args.initial_span*old)
                        bracket=(max(1e-5,old-delta),old+delta)
                    else:
                        bracket=initial
                    warm=best.get(mode)
                    if warm is not None and warm.mesh.mesh_id!=solve_mesh.mesh_id:warm=None
                    if warm is None and mode=='lower' and lower_parent is not None:
                        budget.check('prolonged_lower_original_audit')
                        lower_field=stable_prolong(lower_parent,solve_mesh,best[mode].field,'lower')
                        transported=evaluate(Prepared(solve_mesh,best[mode].fs,mode,8),lower_field)
                        row['lower_prolongation_audit']=transported
                        if transported['witness']:
                            warm=Certificate(mode,best[mode].fs,solve_mesh,lower_field,transported,
                                'independently audited local-coordinate lower prolongation',False)
                    if warm is None and mode=='upper' and mode in best and velocity is not None:
                        budget.check('prolonged_upper_original_audit')
                        transported=evaluate(Prepared(solve_mesh,best[mode].fs,mode,8),velocity)
                        if transported['witness']:
                            warm=Certificate(mode,best[mode].fs,solve_mesh,velocity.copy(),transported,
                                'independently audited prolonged upper velocity',False)
                    candidate,summary=search(solve_mesh,mode,bracket,folder/label/mode,
                        label+'_'+mode,args.tolerance,args.solve_seconds,warm,budget=budget,
                        bisection=args.root_method=='bisection')
                    row[mode]=summary
                    if mode=='upper' and candidate is not None:velocity=candidate.field
                    if candidate is not None and (mode not in best or
                            (candidate.fs>best[mode].fs if mode=='lower' else candidate.fs<best[mode].fs) or
                            (candidate.fs==best[mode].fs and candidate.root_complete and not best[mode].root_complete)):
                        if mode in best and candidate.mesh.physical_model_id!=best[mode].mesh.physical_model_id:
                            raise ValueError('Physical models differ')
                        best[mode]=candidate
                    report=checkpoint()
                    if report['target_met']:
                        return checkpoint('FINISHED','target_met_after_'+mode)
                    budget.check('completed_'+label+'_'+mode)
            report=checkpoint()
            if set(best)=={'lower','upper'}:
                print('BOUNDS',args.case,len(mesh.tri),best['lower'].fs,best['upper'].fs,
                      report['relative_width_symmetric'],flush=True)
                if report['target_met']:
                    stop_reason='target_met';break
            if velocity is None or len(velocity)<12*len(mesh.tri):
                stop_reason='no_usable_upper_velocity';break
        return checkpoint('FINISHED',stop_reason)
    except BudgetExceeded as exc:
        return checkpoint('BUDGET_EXHAUSTED',str(exc))
    except Exception as exc:
        checkpoint('ERROR',type(exc).__name__+': '+str(exc))
        raise


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--case', required=True)
    parser.add_argument('--seed', type=Path)
    parser.add_argument('--resume', action='store_true')
    parser.add_argument('--recheck', action='store_true',help='Re-solve each retained side on its own witness mesh first')
    parser.add_argument('--fan-lower',action='store_true',help='One angle-space split on each proposed lower mesh; retain independent witnesses')
    parser.add_argument('--upper-only',action='store_true',help='Retain the independent lower witness while refining only the upper mesh after recheck')
    parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--cycles', type=int, default=8)
    parser.add_argument('--max-elements', type=int, default=12000)
    parser.add_argument('--theta', type=float, default=.6)
    parser.add_argument('--lower-theta',type=float,default=.8)
    parser.add_argument('--independent-lower',action='store_true')
    parser.add_argument('--upper-indicator',choices=['mixed','kinematic'],default='mixed')
    parser.add_argument('--kinematic-fraction',type=float,default=.2)
    parser.add_argument('--initial-span',type=float,default=.005)
    parser.add_argument('--root-method',choices=['bisection','secant'],default='bisection')
    parser.add_argument('--initial-mesh',type=Path,help='Fresh mesh-only npz/bin; never imports fields or roots')
    parser.add_argument('--tolerance', type=float, default=1e-4)
    parser.add_argument('--target-gap', type=float, default=.02, help='Symmetric certified-pair target; does not relax root or physical gates')
    parser.add_argument('--case-seconds', type=float, default=900., help='Cooperative whole-case wall budget; supervise in-flight native work externally')
    parser.add_argument('--solve-seconds', type=float, default=30., help='Maximum per solver call, further limited by remaining case budget')
    run(parser.parse_args())
