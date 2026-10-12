"""Independent Python/SOCP/LP diagnosis, same saved native Excel coarse mesh.
Never changes cohesion, Davis policy, Fs definition or any source workbook.
"""
import argparse,collections,json,math,sys,time,copy,os
for key in ['OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS']:os.environ.setdefault(key,'1')
import numpy as np
from scipy import sparse
from scipy.optimize import linprog
import clarabel
from support import ROOT,REPO,BASE,sha,write
from common import Mesh,assemble_lower,assemble,physical_audit,save_mesh
from model_support.schema import topology
from conic_lp import build_facets,cone_error,affine_reduce

def mesh_from_saved():
 p=ROOT/'data/P06N_original243_cells.json';cells=json.loads(p.read_text(encoding='utf8'))['cells']
 rec=json.loads((ROOT/'data/inputs/P06N.json').read_text(encoding='utf8'))
 nc=cells['節点データ'];ec=cells['要素データ']
 nr=sorted(int(k[1:]) for k in nc if k.startswith('A') and k!='A1');er=sorted(int(k[1:]) for k in ec if k.startswith('A') and k!='A1')
 assert [int(nc['A'+str(i)]) for i in nr]==list(range(1,len(nr)+1))
 xy=np.asarray([[nc['B'+str(i)],nc['C'+str(i)]] for i in nr]);tri=np.asarray([[int(ec[c+str(i)])-1 for c in ['B','C','D']] for i in er],dtype=np.int32)
 ids=np.asarray([int(ec['E'+str(i)])-1 for i in er],dtype=np.int32)
 assert len(xy)==146 and len(tri)==243
 return mesh_from_arrays(xy,tri,ids,rec),rec

def mesh_from_arrays(xy,tri,ids,rec):
 mats=np.asarray([[m['c'],m['phi'],m['psi'],m['gamma']] for m in rec['source_case']['materials']])
 counts=collections.Counter(tuple(sorted((int(a),int(b)))) for t in tri for a,b in zip(t,np.roll(t,-1)))
 bd={};kinds={'自由':'load','固定XY':'fixed','固定X':'roller_x','固定Y':'roller_y'}
 def on(p,a,b):
  a=np.asarray(a);b=np.asarray(b);v=b-a;t=(p-a)@v/(v@v)
  return -1e-10<=t<=1+1e-10 and abs(v[0]*(p[1]-a[1])-v[1]*(p[0]-a[0]))<1e-9*np.linalg.norm(v)
 for edge,num in counts.items():
  if num!=1:continue
  matches=[f for f in rec['boundary_explicit'] if all(on(xy[i],f['a'],f['b']) for i in edge)];assert len(matches)==1,(edge,matches)
  f=matches[0];loads=np.zeros((2,2));loads[:,1]=np.asarray(f['load0'])+np.asarray(f['load1'])
  bd[edge]={'kind':kinds[f['kind']],'rigid':-1,'loads':loads}
 body=np.zeros((len(tri),2,2));body[:,1,1]=-mats[ids,3]
 return Mesh(xy,tri,mats,ids,np.full(len(tri),-1),body,bd,max(1.,float(mats[:,0].max())),sha(ROOT/'data/inputs'/(rec['case']+'.json')),provenance={'psi_policy':rec['source_case']['policy'],'case':rec['case']}).validate()

def solve(m,fs,mode,fixed_load=None,seconds=30,reduce=False,settings_overrides=None):
 if mode=='lower':
  E,b,G,h,dims,zero=assemble_lower(m,fs,fixed_load);cost=np.zeros(E.shape[1]);cost[-1]=-1. if fixed_load is None else 0.
 else:
  p=assemble(m,fs,8);E,b,G,h,dims,cost=p.eq,p.rhs,p.G,np.zeros(p.G.shape[0]),p.dims,p.objective;zero=None
 original=(E,b,G,h,dims,cost,zero)
 offset=np.zeros(len(cost));T=None;stats=[]
 if reduce:
  E,b,G,h,dims,T,offset,stats=affine_reduce(E,b,G,h,dims)
  reduced_cost=np.asarray(cost@T).ravel()
 else:reduced_cost=cost
 A=sparse.vstack([E,-G],format='csc');rhs=np.r_[b,h];cones=[clarabel.ZeroConeT(len(b))]+[clarabel.NonnegativeConeT(d) if d==1 else clarabel.SecondOrderConeT(d) for d in dims]
 settings=clarabel.DefaultSettings();settings.verbose=False;settings.max_iter=250;settings.time_limit=seconds
 settings.tol_gap_abs=settings.tol_gap_rel=settings.tol_feas=1e-10
 settings.tol_infeas_abs=settings.tol_infeas_rel=1e-10
 settings.max_threads=1
 for key,value in (settings_overrides or {}).items():setattr(settings,key,value)
 solver=clarabel.DefaultSolver(sparse.csc_matrix((len(reduced_cost),len(reduced_cost))),reduced_cost,A,rhs,cones,settings);s=solver.solve()
 x=np.asarray(s.x);z=np.asarray(s.z)
 if T is not None:x=np.asarray(T@x+offset).ravel()
 E0,b0,G0,h0,dims0,cost0,zero=original;v=G0@x+h0
 r={'fs':fs,'mode':mode,'fixed_load':fixed_load,'status':str(s.status),'objective':float(cost0@x),'load':float(x[-1]) if mode=='lower' else None,'seconds':s.solve_time,'iterations':s.iterations,'eq_error':float(np.max(np.abs(E0@x-b0),initial=0)),'cone_error':cone_error(v,dims0),'zero_stress_controls':int(zero.sum()) if zero is not None else None,'solver_primal_obj':s.obj_val,'solver_dual_obj':s.obj_val_dual,'dual_gap':abs(s.obj_val-s.obj_val_dual),'reduced':reduce,'reduction_stats':stats}
 r['dual_stationarity_max']=float(np.max(np.abs(reduced_cost+A.T@z),initial=0))
 r['dual_cone_error']=cone_error(z[len(b):],dims)
 if str(s.status) in ['Solved','AlmostSolved']:r['physical_audit']=physical_audit(m,x,mode,fs,8)
 if str(s.status) in ['PrimalInfeasible','AlmostPrimalInfeasible']:
  r['infeasibility_certificate']={'stationarity_max':float(np.max(np.abs(A.T@z),initial=0)),'rhs_dot_dual':float(rhs@z),'SOC_dual_error':cone_error(z[len(b):],dims)}
 return r,x,original,z

def outer_lp(pr):
 E,b,G,h,dims,cost,zero=pr;A,rhs,heads=build_facets(G,h,dims,8,False)
 t=time.perf_counter();s=linprog(cost,A_ub=A,b_ub=rhs,A_eq=E,b_eq=b,bounds=(None,None),method='highs',options={'time_limit':30,'primal_feasibility_tolerance':1e-9,'dual_feasibility_tolerance':1e-9})
 r={'status':int(s.status),'message':s.message,'seconds':time.perf_counter()-t}
 if s.x is not None:
  r.update(load=float(s.x[-1]),objective=float(s.fun),equality_error=float(np.max(np.abs(E@s.x-b),initial=0)),inequality_error=float(np.max(A@s.x-rhs,initial=0)))
 if s.status==0:
  y=s.eqlin.marginals;w=-s.ineqlin.marginals;dual_obj=float(b@y-rhs@w)
  r.update(dual_objective=dual_obj,dual_stationarity_max=float(np.max(np.abs(cost-E.T@y+A.T@w),initial=0)),nonnegative_weights=bool(w.min(initial=0)>=-1e-12),upper_bound_on_true_lower_lambda=-dual_obj)
 return r,s

def main():
 p=argparse.ArgumentParser();p.add_argument('--tag',default='initial_python');args=p.parse_args();folder=ROOT/'results'/args.tag;folder.mkdir(exist_ok=False)
 m,rec=mesh_from_saved();save_mesh(folder/'exact_mesh.npz',m)
 area,grad,edges=topology(m);regions=np.bincount(m.material_id,weights=area,minlength=3)
 records={'python_version':sys.version,'clarabel_version':clarabel.__version__,'mesh_id':m.mesh_id,'nodes':len(m.xy),'elements':len(m.tri),'region_areas':regions.tolist(),'total_weight':float(area@m.materials[m.material_id,3]),'policy':m.provenance,'scope':'Same actual243-element coarse mesh; no full adaptive run','fixed_points':[]}
 assert np.max(np.abs(regions-np.asarray(rec['region_areas'])))<1e-9
 for fs,mode,fixed in [(1.,'upper',None),(1.,'lower',None),(.5,'lower',None),(.125,'lower',None),(.5,'lower',1.)]:
  r,x,pr,z=solve(m,fs,mode,fixed);np.savez_compressed(folder/(mode+'_'+str(fs)+'_'+str(fixed)+'.npz'),x=x,z=z)
  if mode=='lower' and fixed is None:r['outer_polygon_LP'],unused=outer_lp(pr)
  records['fixed_points'].append(r);write(folder/'summary.json',records);print(json.dumps(r,ensure_ascii=False),flush=True)
 # Which zero-control elements/fans are implicated? Preserve the full coordinates.
 _,_,_,_,_,zero=assemble_lower(m,.5)
 records['zero_controls_by_element']=[{'element':e,'material':int(m.material_id[e])+1,'vertices':m.xy[m.tri[e]].tolist(),'zero':np.flatnonzero(zero[e]).tolist()} for e in range(len(m.tri)) if zero[e].any()]
 write(folder/'summary.json',records)
 print('PYTHON DIAGNOSIS COMPLETE',flush=True)
if __name__=='__main__':main()
