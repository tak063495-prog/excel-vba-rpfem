"""One current original-dual recenter attempt, then strict cached root audit."""
from pathlib import Path
import argparse,hashlib,json,sys,time,shutil
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'python'))
import accuracy_driver as d
import experiment_core as c
import original_dual as od
from upper_proximal_candidate import program_hash
np,sparse=c.np,c.sparse
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()

def run(a):
    started=time.perf_counter()
    if a.output.exists():raise ValueError('Fresh output required')
    budget=d.CaseBudget(a.case_seconds,a.solve_seconds)
    sources=d.source_hashes();sources[str(Path(__file__).resolve())]=sha(__file__)
    _,rec=d.load_model(a.case);mesh=d.load_saved_mesh(a.geometry,rec)
    paths=[a.geometry,a.dual,a.primal,ROOT/'INPUTS_26.json',d.CATALOG/(a.case+'.json')]
    inputs={str(p.resolve()):sha(p) for p in paths}
    with np.load(a.dual,allow_pickle=False) as data:
        expected=dict(case=a.case,mode='upper',mesh_id=mesh.mesh_id,physical_model_id=mesh.physical_model_id,policy=mesh.provenance['psi_policy'])
        if any(str(data[k])!=v for k,v in expected.items()) or data['q'].shape!=() or data['q'].item()!=8:raise ValueError('Original dual identity mismatch')
        fs=float(data['fs']);center=data['dual'].copy()
    with np.load(a.primal,allow_pickle=False) as data:
        if str(data['mode'])!='upper' or data['q'].shape!=() or data['q'].item()!=8:raise ValueError('Original primal identity mismatch')
        upper_fs=float(data['fs']);field=data['x'].copy()
    if not d.finite(fs,upper_fs) or not 0<fs<upper_fs or upper_fs-fs>1e-4*max(1.,(fs+upper_fs)/2):raise ValueError('Unchanged normalized root width required')
    primal=c.Prepared(mesh,upper_fs,'upper',8,reduction='none');E,b,G,h,dims,cost=primal.original
    rows=np.flatnonzero(b!=0)
    if len(rows)!=1 or np.any(h!=0) or not np.isfinite(field).all():raise ValueError('Original homogeneous work layout required')
    work=float((E[int(rows[0])]@field)[0]/b[int(rows[0])])
    if not np.isfinite(work) or work<=0:raise ValueError('Positive work required')
    field/=work;pa=c.evaluate(primal,field)
    if not pa['physical_pass'] or not pa['witness']:raise ValueError('Fresh original primal audit rejected')
    pr=c.Prepared(mesh,fs,'upper',8,reduction='none');E,b,G,h,dims,cost=pr.original
    if od.eligibility(pr):raise ValueError('Original dual program required')
    ny,nz=len(b),len(h);nv=ny+nz
    if center.shape!=(nv,) or not np.isfinite(center).all():raise ValueError('Finite full original dual center required')
    a.output.mkdir(parents=True);copies=a.output/'source_inputs';copies.mkdir()
    for p in [a.geometry,a.dual,a.primal]:
        dest=copies/p.name;shutil.copyfile(p,dest)
        if sha(dest)!=inputs[str(p.resolve())]:raise ValueError('Snapshot changed')
    before_hash=program_hash(pr.original);before=od.audit_original_dual(pr,center)
    stationary=sparse.hstack([E.T,-G.T],format='csc')
    conic=sparse.hstack([sparse.csc_matrix((nz,ny)),-sparse.eye(nz,format='csc')],format='csc')
    matrix=sparse.vstack([stationary,conic,sparse.csc_matrix(np.r_[b,h].reshape(1,-1))],format='csc')
    rhs=np.r_[-cost,np.zeros(nz),-od.threshold(pr)]
    cones=[c.clarabel.ZeroConeT(len(cost))]+[c.clarabel.NonnegativeConeT(k) if k==1 else c.clarabel.SecondOrderConeT(k) for k in dims]+[c.clarabel.NonnegativeConeT(1)]
    config=dict(equilibrate_enable=False,time_limit=budget.allowance('one_original_dual_recenter'),
        tol_gap_abs=1e-13,tol_gap_rel=1e-13,tol_feas=1e-13,iterative_refinement_max_iter=30,
        iterative_refinement_abstol=1e-15,iterative_refinement_reltol=1e-15)
    solver=c.clarabel.DefaultSolver(sparse.eye(nv,format='csc')*.01,-.01*center,matrix,rhs,cones,c.settings(config,350))
    tick=time.perf_counter();solution=solver.solve();solve_seconds=time.perf_counter()-tick
    vector=np.asarray(solution.x).copy()  # helper solution.z is never read.
    audit=od.audit_original_dual(c.Prepared(mesh,fs,'upper',8,reduction='none'),vector)
    identity=dict(case=a.case,mode='upper',q=8,fs=fs,mesh_id=mesh.mesh_id,physical_model_id=mesh.physical_model_id,policy=mesh.provenance['psi_policy'])
    np.savez_compressed(a.output/'recentered_original_dual.npz',dual=vector,**identity)
    warm=d.Certificate('upper',upper_fs,mesh,field,pa,'Current-session raw original QP normalized and fully reaudited',False)
    best=warm;root=None
    if audit['accepted']:
        guide=dict(identity,dual=vector)
        best,root=d.search(mesh,'upper',(fs,upper_fs),a.output/'closed_root',a.case+'_recentered_original_dual_closure',
            1e-4,a.solve_seconds,initial_witness=warm,initial_dual_guides=[guide],budget=budget,bisection=True)
    current=d.source_hashes();current[str(Path(__file__).resolve())]=sha(__file__)
    if current!=sources or any(sha(p)!=h for p,h in inputs.items()) or program_hash(pr.original)!=before_hash:raise ValueError('Frozen source or original program changed')
    result=dict(case=a.case,execution_scope=dict(kind='current_session_original_dual_recenter_closure',historical_field_imported=False,historical_fs_imported=False),
        status='OWN_ROOT_COMPLETED' if best and best.root_complete else 'ORIGINAL_DUAL_REJECTED',target_met=False,
        upper=best.publish(a.output),root=root,root_complete=bool(best.root_complete),
        source_original_dual_audit=before,fresh_original_dual_audit=audit,
        auxiliary_objective='0.5*.01*||original_dual-center||^2',auxiliary_constraint_dual_used=False,
        original_constraints_unchanged=True,original_program_sha256_before=before_hash,original_program_sha256_after=program_hash(pr.original),
        original_physical_primal_audit=pa,positive_work_normalization=work,settings=config,solver_status_metadata=str(solution.status),
        solver_seconds=solve_seconds,seconds=time.perf_counter()-started,case_budget=budget.record(),
        input_sha256=inputs,source_sha256_before=sources,source_sha256_after=current,source_unchanged=True)
    d.write(a.output/'summary.json',result)
    print('RECENTER_CLOSED',a.case,audit,best.root_complete,result['seconds'],flush=True)
    return result
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--case',required=True);p.add_argument('--geometry',type=Path,required=True)
    p.add_argument('--dual',type=Path,required=True);p.add_argument('--primal',type=Path,required=True);p.add_argument('--output',type=Path,required=True)
    p.add_argument('--case-seconds',type=float,default=90.);p.add_argument('--solve-seconds',type=float,default=60.)
    r=run(p.parse_args());raise SystemExit(0 if r['root_complete'] else 2)
