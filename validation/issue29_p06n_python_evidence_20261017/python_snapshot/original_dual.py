"""Solve original dual feasibility solely to guide the discrete root search.

The original physical primal field is never modified or published. All original
stationarity, product-cone and separation constraints are audited independently
of the auxiliary solver's status. Lower programs require the exact free-load
multiplier objective and original field layout.
"""
from __future__ import annotations
import time
from experiment_core import np,sparse,clarabel,settings,cone_error,UPPER_TARGET,LOWER_TARGET


def eligibility(prepared):
    if prepared.mode not in ['lower','upper']:return 'LOWER_OR_UPPER_ONLY'
    if prepared.q!=8:return 'ORIGINAL_Q8_ONLY'
    E,b,G,h,dims,cost=prepared.original
    if not all(np.isfinite(v).all() for v in (E.data,b,G.data,h,cost)):
        return 'NONFINITE_ORIGINAL_PROGRAM'
    if E.shape!=(len(b),len(cost)) or G.shape!=(len(h),len(cost)) or sum(dims)!=len(h) or any(d<=0 for d in dims):
        return 'ORIGINAL_PROGRAM_DIMENSIONS'
    if prepared.mode=='lower':
        expected=np.zeros(18*len(prepared.m.tri)+1);expected[-1]=-1.
        if not np.array_equal(cost,expected):return 'LOWER_FREE_LOAD_OBJECTIVE_AND_LAYOUT_REQUIRED'
    return None


def threshold(prepared):
    return UPPER_TARGET+4e-6 if prepared.mode=='upper' else -LOWER_TARGET+3e-6


def audit_original_dual(prepared,dual):
    reason=eligibility(prepared)
    if reason:return {'accepted':False,'reason':reason}
    E,b,G,h,dims,cost=prepared.original;dual=np.asarray(dual)
    if dual.ndim!=1 or len(dual)!=len(b)+len(h) or not np.isfinite(dual).all():
        return {'accepted':False,'reason':'NONFINITE_OR_WRONG_SIZE_DUAL'}
    y,z=dual[:len(b)],dual[len(b):]
    stat=float(np.max(np.abs(cost+E.T@y-G.T@z),initial=0.))
    cone=cone_error(z,dims);objective=-float(b@y+h@z);required=threshold(prepared)
    accepted=bool(np.isfinite([stat,cone,objective]).all() and stat<=1e-7 and cone<=1e-8 and objective>=required)
    result={'accepted':accepted,'dual_stationarity':stat,'dual_cone_error':cone,
        'dual_objective':objective,'threshold':required,'method':'original_dual_feasibility',
        'scope':'discrete_root_search_only','publishes_primal_bound':False}
    if accepted:result['search_sign']=1 if prepared.mode=='upper' else -1
    else:result['reason']='ORIGINAL_RAW_DUAL_GATE_REJECTED'
    return result


def solve_original_dual(prepared,seconds=20.,return_dual=False,*,equilibrate=True,regularization=0.,equality_scale=False):
    """Return raw audited original evidence; optionally include its full dual."""
    started=time.perf_counter();reason=eligibility(prepared)
    if reason:return {'accepted':False,'reason':reason,'method':'original_dual_feasibility','seconds':0.}
    if not np.isfinite(seconds) or seconds<=0:
        return {'accepted':False,'reason':'INVALID_TIME_LIMIT','method':'original_dual_feasibility','seconds':0.}
    if not np.isfinite(regularization) or regularization<0:
        return {'accepted':False,'reason':'INVALID_AUXILIARY_REGULARIZATION','method':'original_dual_feasibility','seconds':0.}
    E,b,G,h,dims,cost=prepared.original;ny,nz=len(b),len(h);nv=ny+nz
    stationary=sparse.hstack([E.T,-G.T],format='csc')
    conic=sparse.hstack([sparse.csc_matrix((nz,ny)),-sparse.eye(nz,format='csc')],format='csc')
    separation=sparse.csc_matrix(np.r_[b,h].reshape(1,-1))
    A=sparse.vstack([stationary,conic,separation],format='csc')
    rhs=np.r_[-cost,np.zeros(nz),-threshold(prepared)]
    cones=[clarabel.ZeroConeT(len(cost))]+[clarabel.NonnegativeConeT(d) if d==1 else clarabel.SecondOrderConeT(d) for d in dims]+[clarabel.NonnegativeConeT(1)]
    if equality_scale:
        factor=np.asarray(abs(stationary).max(axis=1).toarray()).ravel()
        factor=1./np.maximum(1.,factor)
        A=sparse.diags(np.r_[factor,np.ones(nz+1)])@A;rhs[:len(cost)]*=factor
    config={'equilibrate_enable':bool(equilibrate),'time_limit':float(seconds),
            'tol_gap_abs':1e-13,'tol_gap_rel':1e-13,'tol_feas':1e-13,
            'iterative_refinement_max_iter':30,'iterative_refinement_abstol':1e-15,'iterative_refinement_reltol':1e-15}
    build=time.perf_counter()-started
    try:
        P=sparse.eye(nv,format='csc')*regularization if regularization else sparse.csc_matrix((nv,nv))
        solver=clarabel.DefaultSolver(P,np.zeros(nv),A.tocsc(),rhs,cones,settings(config,350))
        solver.solve();solution=solver.get_solution();dual=np.asarray(solution.x)
        evidence=audit_original_dual(prepared,dual)
        evidence.update(status=str(solution.status),iterations=solution.iterations,settings=config,max_iter=350,
            build_seconds=build,seconds=time.perf_counter()-started,fs=prepared.fs,mode=prepared.mode,q=8,
            mesh_id=prepared.m.mesh_id,physical_model_id=prepared.m.physical_model_id,case=prepared.m.provenance['case'])
        evidence.update(auxiliary_regularization=regularization,equality_scale=bool(equality_scale))
        if return_dual:evidence['dual']=dual
        return evidence
    except Exception as error:
        return {'accepted':False,'reason':f'SOLVER_EXCEPTION:{type(error).__name__}:{error}',
            'method':'original_dual_feasibility','settings':config,'seconds':time.perf_counter()-started}
