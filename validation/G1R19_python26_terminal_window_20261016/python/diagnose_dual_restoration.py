"""Scratch equality-dual restoration at an UNKNOWN original zero-cost upper point."""
from pathlib import Path
import time
import math
import hashlib
from scipy.linalg import lstsq
from scipy.sparse.linalg import lsmr
import experiment_core as core
import accuracy_driver as driver
np=core.np


def main():
    root=Path(__file__).resolve().parents[1]
    case,fs='D01A',.579209921848832
    _,rec=driver.load_model(case)
    path=root/'results/d01_recovery/D01A/D01A_C2_mesh.npz'
    mesh=driver.load_saved_mesh(path,rec)
    pr=core.Prepared(mesh,fs,'upper',8,reduction='affine')
    E,b,G,h,dims,cost=pr.original
    assert np.all(cost==0)
    EE=(E@pr.T).tocsr();EE.eliminate_zeros()
    active=np.flatnonzero(np.diff(EE.indptr)>0)
    removed=np.setdiff1d(np.arange(len(b)),active)
    M=E[removed].T.tocsc()
    output=root/'results'/('dual_restore_diagnostic_'+str(time.time_ns())+'.json')
    report={'case':case,'fs':fs,'q':8,'mesh_id':mesh.mesh_id,'mesh_path':str(path),
        'original_E_shape':E.shape,'removed_M_shape':M.shape,'source_sha256':{
        name:hashlib.sha256(Path(__file__).with_name(name).read_bytes()).hexdigest()
        for name in ('experiment_core.py','accuracy_driver.py','reduction.py')},'configurations':[]}
    def metrics(z):
        y,zc=z[:len(b)],z[len(b):]
        return {'objective':-float(b@y+h@zc),
            'stationarity':float(np.max(np.abs(E.T@y-G.T@zc),initial=0)),
            'stationarity_norm':float(np.linalg.norm(E.T@y-G.T@zc)),
            'cone_error':core.cone_error(zc,dims)}
    def candidate(label,z,timer,solver_info=None):
        before=metrics(z);obj=before['objective']
        factor=(core.UPPER_TARGET+4e-6)/obj if obj>0 else None
        after=metrics(factor*z) if factor and math.isfinite(factor) and factor>0 else None
        passed=bool(after and all(math.isfinite(v) for v in after.values()) and
            after['stationarity']<=1e-7 and after['cone_error']<=1e-8 and
            after['objective']>=core.UPPER_TARGET+2e-6)
        row={'method':label,'before':before,'factor':factor,'after':after,
            'raw_audit_pass':passed,'seconds':time.perf_counter()-timer,'solver':solver_info}
        print(label,before,'scaled_raw_pass',passed,flush=True)
        return row
    for scaled in (False,True):
        options=dict(driver.ACCURATE,iterative_refinement_max_iter=30,
            iterative_refinement_abstol=1e-15,iterative_refinement_reltol=1e-15,time_limit=120)
        if scaled:options.update(tol_gap_abs=1e-13,tol_gap_rel=1e-13,tol_feas=1e-13)
        solve_result,_,canonical=core.solve(pr,options,max_iter=350,cone_scaling=scaled,primal_polish=True)
        z0,rest=pr.restore_dual(canonical)
        current={'scaled':scaled,'status':solve_result['status'],'default_restoration':rest,'candidates':[]}
        report['configurations'].append(current)
        current['candidates'].append(candidate('existing_restore',z0,time.perf_counter()))
        base=np.r_[np.zeros(len(b)),z0[len(b):]]
        base[active]=canonical[:len(pr.reduced[1])]
        rhs=-(E.T@base[:len(b)]-G.T@base[len(b):])
        t=time.perf_counter();result=lsmr(M,rhs,atol=1e-15,btol=1e-15,conlim=1e16,maxiter=100000)
        z=base.copy();z[removed]=result[0]
        current['candidates'].append(candidate('removed_lsmr_1e15_conlim1e16_100k',z,t,
            {'flag':result[1],'iterations':result[2],'normr':result[3],'normar':result[4],'normA':result[5],'condition':result[6]}))
        if M.shape[0]*M.shape[1]<=10_000_000:
            t=time.perf_counter();dense=M.toarray()
            solution,residual,rank,_=lstsq(dense,rhs,cond=1e-15,lapack_driver='gelsy',overwrite_a=True)
            z=base.copy();z[removed]=solution
            current['candidates'].append(candidate('removed_dense_gelsy',z,t,{'rank':rank}))
        else:current['dense_QR_skipped']='removed matrix exceeds 80MB double allocation cap'
        t=time.perf_counter();residual=E.T@z0[:len(b)]-G.T@z0[len(b):]
        correction=lsmr(E.T.tocsc(),-residual,atol=1e-15,btol=1e-15,conlim=1e16,maxiter=100000)
        z=z0.copy();z[:len(b)]+=correction[0]
        current['candidates'].append(candidate('full_original_equality_lsmr_correction',z,t,
            {'flag':correction[1],'iterations':correction[2],'normr':correction[3],
             'normar':correction[4],'normA':correction[5],'condition':correction[6]}))
        core.write(output,report)
    print('DIAGNOSTIC',output,flush=True)


if __name__=='__main__':main()
