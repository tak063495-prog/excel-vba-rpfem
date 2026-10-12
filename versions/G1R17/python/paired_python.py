"""P06A/N original/repaired same coarse geometry, independent SOCP + LP."""
import json,numpy as np
from scipy import sparse
from scipy.optimize import linprog
from support import ROOT,write
from conic_lp import build_facets
from diagnose import mesh_from_saved,mesh_from_arrays,solve,outer_lp
from fan_repair import repair
def localized_dual(pr):
 E,b,G,h,dims,q,zero=pr;A,rhs,unused=build_facets(G,h,dims,8,False)
 support=json.loads((ROOT/'results/localization.json').read_text(encoding='utf8'))
 ei=[r['row']for r in support['active_equalities']];ai=[r['row']for r in support['active_cones']]
 B=sparse.hstack([E[ei].T,-A[ai].T],format='csc')
 s=linprog(np.zeros(len(ei)+len(ai)),A_eq=B,b_eq=q,bounds=[(None,None)]*len(ei)+[(0,None)]*len(ai),method='highs',options={'time_limit':10})
 assert s.status==0,s.message
 y=s.x[:len(ei)];w=s.x[len(ei):];residual=float(np.max(abs(B@s.x-q)));dual=float(b[ei]@y-rhs[ai]@w)
 assert residual<1e-10 and w.min()>=0 and dual==0
 return {'status':0,'dual_stationarity_max':residual,'dual_objective':dual,'nonnegative_weights':True,'upper_bound_on_true_lower_lambda':-dual,'eq_indices':ei,'ineq_indices':ai,'equality_dual':y,'nonnegative_dual':w,'scope':'Localized support; certificate verified against every original variable'}
def main():
 old,_=mesh_from_saved();out={'cases':[]}
 for case in ['P06N','P06A']:
  rec=json.loads((ROOT/'data/inputs'/(case+'.json')).read_text(encoding='utf8'))
  m=mesh_from_arrays(old.xy.copy(),old.tri.copy(),old.material_id.copy(),rec)
  repaired,changes=repair(m,rec);entry={'case':case,'original_mesh':m.mesh_id,'repaired_mesh':repaired.mesh_id,'changes':changes,'points':[]};out['cases'].append(entry)
  for mesh,label in [(m,'original243'),(repaired,'repaired245')]:
   for fs in [.5,1.]:
    r,x,p,z=solve(mesh,fs,'lower');r['mesh']=label
    if label=='original243':r['outer_polygon_LP_dual']=localized_dual(p)
    else:assert r.get('physical_audit',{}).get('pass'),r
    entry['points'].append(r);print(case,label,fs,r['status'],r['load'],flush=True)
    write(ROOT/'results/paired_python.json',out)
  # Native lower endpoint and opposite endpoint, only N known initially.
  if case=='P06N':
   for fs in [.5601486798687134,.560219483290432]:
    r,x,p,z=solve(repaired,fs,'lower',fixed_load=1.)
    r['mesh']='repaired245_fixed_lambda1';entry['points'].append(r);print(case,'lambda1',fs,r['status'],r['eq_error'],r['cone_error'],flush=True)
    write(ROOT/'results/paired_python.json',out)
 write(ROOT/'results/paired_python.json',out)
if __name__=='__main__':main()
