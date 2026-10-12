"""Localize the independent LP dual and separate equalities from yield constraints."""
import json,numpy as np
from scipy.optimize import linprog
from support import ROOT,write
from diagnose import mesh_from_saved,solve,outer_lp
from model_support.schema import topology

def main():
 m,rec=mesh_from_saved();r,x,pr,z=solve(m,.5,'lower');E,b,G,h,dims,cost,zero=pr
 equalonly=linprog(cost,A_eq=E,b_eq=b,bounds=[(None,None)]*(len(cost)-1)+[(0,1)],method='highs',options={'time_limit':30})
 lp,s=outer_lp(pr);assert s.status==0
 area,grad,edges=topology(m);roles=[]
 for e in range(len(m.tri)):
  for k in range(3):
   for j in range(2):roles.append({'kind':'SOIL_EQUILIBRIUM','element':e,'vertex':k,'component':j})
  for k in range(6):
   if zero[e,k]:
    for j in range(3):roles.append({'kind':'ZERO_STRESS','element':e,'control':k,'component':j})
 for ed in edges:
  for k in range(3):
   for j in range(2):
    if ed.elements[1]>=0 or ed.kind=='load' or (ed.kind=='roller_x' and j==1) or(ed.kind=='roller_y' and j==0):roles.append({'kind':'INTERNAL_TRACTION' if ed.elements[1]>=0 else 'BOUNDARY_TRACTION','elements':ed.elements,'control':k,'component':j,'nodes':[ed.a,ed.b],'positions':m.xy[[ed.a,ed.b]].tolist()})
 assert len(roles)==E.shape[0],(len(roles),E.shape)
 cr=[]
 for e in range(len(m.tri)):
  for k in range(6):
   if not zero[e,k]:cr.append({'element':e,'control':k,'material':int(m.material_id[e])+1,'vertices':m.xy[m.tri[e]].tolist()})
 active_eq=[{'row':i,'dual':float(v),**roles[i]} for i,v in enumerate(s.eqlin.marginals) if abs(v)>1e-8]
 nh=len(cr);active_cones=[{'row':i,'dual':float(v),'facet':int(i//nh)-1,**cr[i%nh]} for i,v in enumerate(s.ineqlin.marginals) if abs(v)>1e-8]
 involved=sorted(set([r['element'] for r in active_eq if 'element'in r]+[r['element'] for r in active_cones]))
 result={'fs':.5,'mesh_id':m.mesh_id,'LP_certificate':lp,'equalities_only_with_lambda_cap':{'status':int(equalonly.status),'load':float(equalonly.x[-1]) if equalonly.x is not None else None},'active_equalities':active_eq,'active_cones':active_cones,'involved_elements':[{'id':e,'material':int(m.material_id[e])+1,'vertices':m.xy[m.tri[e]].tolist(),'zero_controls':np.flatnonzero(zero[e]).tolist()}for e in involved]}
 np.savez_compressed(ROOT/'results/initial_python/dual_certificate.npz',eq_dual=s.eqlin.marginals,inequality_dual=s.ineqlin.marginals)
 write(ROOT/'results/localization.json',result)
 print('EQUALITY_ONLY',result['equalities_only_with_lambda_cap'],'ACTIVE_EQ',len(active_eq),'ACTIVE_CONES',len(active_cones),'ELEMENTS',involved,flush=True)
 print(json.dumps(result['involved_elements']),flush=True)
if __name__=='__main__':main()
