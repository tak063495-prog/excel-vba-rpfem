"""Minimum-change feasibility with an objective cap, never a dual root guide."""
import os,sys,argparse,json,time
from pathlib import Path
for n in ['OPENBLAS_NUM_THREADS','MKL_NUM_THREADS','OMP_NUM_THREADS']:os.environ[n]='1'
sys.path.insert(0,r'D:\Work\excel-vba-rpfem\validation\accuracy_20261013\python')
import accuracy_driver as driver
import experiment_core as core
np,sparse,clarabel=core.np,core.sparse,core.clarabel
p=argparse.ArgumentParser();p.add_argument('--case',required=True);p.add_argument('--seed',type=Path,required=True)
p.add_argument('--output',type=Path,required=True);p.add_argument('--alpha',type=float,default=.01)
a=p.parse_args()
if a.output.exists() or not np.isfinite(a.alpha) or a.alpha<=0:raise ValueError('New output and positive alpha required')
meta=json.loads((a.seed/'audit.json').read_text(encoding='utf8'));assert meta['case']==a.case and meta['q']==8
_,rec=driver.load_model(a.case);mesh=driver.load_saved_mesh(a.seed/'original_mesh.npz',rec)
assert meta['mesh_id']==mesh.mesh_id and meta['physical_model_id']==mesh.physical_model_id
with np.load(a.seed/'original_candidate.npz',allow_pickle=False) as data:source=data['x'].copy()
fs=meta['fs'];pr=core.Prepared(mesh,fs,'upper',8,reduction='none');E,b,G,h,dims,cost=pr.original
cap=core.UPPER_TARGET-1e-6
matrix=sparse.vstack([E,-G,sparse.csc_matrix(cost.reshape(1,-1))],format='csc')
rhs=np.r_[b,h,cap]
cones=pr.cones+[clarabel.NonnegativeConeT(1)]
config=dict(driver.ACCURATE,time_limit=600.,tol_gap_abs=1e-13,tol_gap_rel=1e-13,tol_feas=1e-13,
    iterative_refinement_max_iter=30,iterative_refinement_abstol=1e-15,iterative_refinement_reltol=1e-15)
a.output.mkdir(parents=True);(a.output/'scratch_source.py').write_bytes(Path(__file__).read_bytes())
driver.save_mesh(a.output/'original_mesh.npz',mesh)
np.savez_compressed(a.output/'source_candidate.npz',x=source,fs=fs,mode='upper',q=8)
solver=clarabel.DefaultSolver(sparse.eye(len(source),format='csc')*a.alpha,-a.alpha*source,
    matrix,rhs,cones,core.settings(config,350))
print('START_UPPER_CAP_QP',a.case,len(mesh.tri),fs,a.alpha,flush=True)
t=time.perf_counter();solution=solver.solve();seconds=time.perf_counter()-t
field=np.asarray(solution.x);modified_dual=np.asarray(solution.z)
audit=core.evaluate(pr,field)
audit.update(case=a.case,mode='upper',q=8,fs=fs,mesh_id=mesh.mesh_id,physical_model_id=mesh.physical_model_id)
np.savez_compressed(a.output/'original_candidate.npz',x=field,modified_dual=modified_dual,fs=fs,mode='upper',q=8)
report={'case':a.case,'fs':fs,'q':8,'mesh_id':mesh.mesh_id,'physical_model_id':mesh.physical_model_id,
    'source':str(a.seed),'source_sha256':driver.SOURCE_HASHES,'config':config,'alpha':a.alpha,'cap':cap,
    'status':str(solution.status),'iterations':solution.iterations,'seconds':seconds,'original_audit':audit,
    'root_complete':False,'modified_dual_is_original_root_guide':False}
if audit['physical_pass'] and audit['witness']:
    cert=driver.Certificate('upper',fs,mesh,field,audit,'capped proximal QP, independently audited original q8',False)
    core.write(a.output/'import_audit.json',{'case':a.case,'certificates':{'upper':cert.publish(a.output)}})
core.write(a.output/'audit.json',report)
print('UPPER_CAP_QP',a.case,solution.status,audit['value'],audit['physical_pass'],audit['witness'],audit['cone_error'],seconds,flush=True)
