"""Original-cone feasibility evidence for discrete zero-cost upper root search.

An accepted vector only guides the discrete search. This helper never changes
or publishes a physical primal field. Every raw original constraint is checked;
solver status and transformed residuals cannot establish a search sign.
"""
from __future__ import annotations
import time
from experiment_core import np,sparse,clarabel,settings,cone_error

TARGET=1.000003


def eligibility(prepared):
    if prepared.mode!='upper':return 'UPPER_ONLY'
    if prepared.q!=8:return 'ORIGINAL_Q8_ONLY'
    E,b,G,h,dims,cost=prepared.original
    if not np.array_equal(cost,np.zeros_like(cost)):return 'ORIGINAL_COST_NOT_EXACTLY_ZERO'
    if not all(np.isfinite(v).all() for v in (E.data,b,G.data,h,cost)):
        return 'NONFINITE_ORIGINAL_PROGRAM'
    if E.shape!=(len(b),len(cost)) or G.shape!=(len(h),len(cost)) or sum(dims)!=len(h) or any(d<=0 for d in dims):
        return 'ORIGINAL_PROGRAM_DIMENSIONS'
    return None


def audit_original_farkas(prepared,dual):
    reason=eligibility(prepared)
    if reason:return {'accepted':False,'reason':reason}
    E,b,G,h,dims,cost=prepared.original;dual=np.asarray(dual)
    if dual.ndim!=1 or len(dual)!=len(b)+len(h) or not np.isfinite(dual).all():
        return {'accepted':False,'reason':'NONFINITE_OR_WRONG_SIZE_DUAL'}
    y,z=dual[:len(b)],dual[len(b):]
    stat=float(np.max(np.abs(E.T@y-G.T@z),initial=0.))
    cone=cone_error(z,dims);objective=-float(b@y+h@z);normalization=abs(objective-TARGET)
    finite=bool(np.isfinite([stat,cone,objective,normalization]).all())
    accepted=bool(finite and stat<=1e-7 and cone<=1e-8 and objective>=1.000001 and normalization<=1e-7)
    result={'accepted':accepted,'dual_stationarity':stat,'dual_cone_error':cone,
        'dual_objective':objective,'normalization_error':normalization,'target':TARGET,
        'method':'original_cone_farkas_feasibility','scope':'discrete_root_search_only',
        'publishes_primal_bound':False}
    if accepted:result['positive_sign']=1
    else:result['reason']='ORIGINAL_RAW_FARKAS_GATE_REJECTED'
    return result


def solve_original_farkas(prepared,seconds=15.,return_dual=False):
    """Return plain audited evidence; optionally include the full original dual."""
    started=time.perf_counter();reason=eligibility(prepared)
    if reason:return {'accepted':False,'reason':reason,'method':'original_cone_farkas_feasibility','seconds':0.}
    if not np.isfinite(seconds) or seconds<=0:
        return {'accepted':False,'reason':'INVALID_TIME_LIMIT','method':'original_cone_farkas_feasibility','seconds':0.}
    E,b,G,h,dims,cost=prepared.original;ny,nz=len(b),len(h);nv=ny+nz
    # Original self-dual nonnegative/SOC blocks stay in their original order.
    # E.T y - G.T z = 0; -(b.y+h.z)=TARGET; z in the original product cone.
    stationary=sparse.hstack([E.T,-G.T],format='csc')
    separation=sparse.csc_matrix(np.r_[-b,-h].reshape(1,-1))
    conic=sparse.hstack([sparse.csc_matrix((nz,ny)),-sparse.eye(nz,format='csc')],format='csc')
    A=sparse.vstack([stationary,separation,conic],format='csc')
    rhs=np.r_[np.zeros(len(cost)),TARGET,np.zeros(nz)]
    cones=[clarabel.ZeroConeT(len(cost)+1)]+[clarabel.NonnegativeConeT(d) if d==1 else clarabel.SecondOrderConeT(d) for d in dims]
    config={'equilibrate_enable':False,'time_limit':float(seconds)}
    build=time.perf_counter()-started
    try:
        solver=clarabel.DefaultSolver(sparse.csc_matrix((nv,nv)),np.zeros(nv),A,rhs,cones,settings(config,350))
        solver.solve();solution=solver.get_solution();dual=np.asarray(solution.x)
        evidence=audit_original_farkas(prepared,dual)
        evidence.update(status=str(solution.status),iterations=solution.iterations,
            settings=config,max_iter=350,build_seconds=build,seconds=time.perf_counter()-started,
            fs=prepared.fs,mode='upper',q=8,mesh_id=prepared.m.mesh_id,
            physical_model_id=prepared.m.physical_model_id,case=prepared.m.provenance['case'])
        if return_dual:evidence['dual']=dual
        return evidence
    except Exception as error:
        return {'accepted':False,'reason':f'SOLVER_EXCEPTION:{type(error).__name__}:{error}',
            'method':'original_cone_farkas_feasibility','settings':config,
            'seconds':time.perf_counter()-started}
