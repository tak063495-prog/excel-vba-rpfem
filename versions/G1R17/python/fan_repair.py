"""Select proof-exposed zero-stress vertex fans by geometry, not case names."""
import json,numpy as np
from support import ROOT,write
from diagnose import mesh_from_saved,mesh_from_arrays,solve
from model_support.schema import topology
from common import assemble_lower

def repair(m,rec):
 _,_,_,_,_,zero=assemble_lower(m,1.)
 candidates=[]
 for e,t in enumerate(m.tri):
  if m.materials[m.material_id[e],0]!=0 or not np.any(m.body[e]):continue
  for k in range(3):
   if not zero[e,k]:continue
   p=m.xy[t[k]];u=m.xy[t[(k+1)%3]]-p;v=m.xy[t[(k+2)%3]]-p
   if u@v < -1e-12*np.linalg.norm(u)*np.linalg.norm(v):
    # Strictly interior point on the angle bisector; both children at the
    # exposed vertex have acute angles even for strongly unequal edge lengths.
    lu=np.linalg.norm(u);lv=np.linalg.norm(v);a=.5*lv/(lu+lv);b=.5*lu/(lu+lv)
    center=p+a*u+b*v
    candidates.append((e,k,center));break
 if not candidates:return m,[]
 xy=m.xy.tolist();tri=m.tri.tolist();ids=m.material_id.tolist()
 for e,k,center in candidates:
  mid=len(xy);xy.append(center.tolist());a,b,c=tri[e]
  tri[e]=[a,b,mid];tri.append([b,c,mid]);tri.append([c,a,mid]);ids.extend([ids[e]]*2)
 return mesh_from_arrays(np.asarray(xy),np.asarray(tri,np.int32),np.asarray(ids,np.int32),rec),[{'element':e,'vertex':k,'point':c.tolist()} for e,k,c in candidates]

def main():
 m,rec=mesh_from_saved();new,changes=repair(m,rec);unchanged,repeat=repair(new,rec);assert not repeat
 out={'nodes':len(new.xy),'elements':len(new.tri),'changes':changes,'repeat_changes':repeat,'fixed_points':[]}
 assert len(changes)==1 and changes[0]['element']==65 and len(new.tri)==245
 for fs in [1.,.65,.5,.125]:
  r,x,pr,z=solve(new,fs,'lower');assert r.get('physical_audit',{}).get('pass'),r
  out['fixed_points'].append(r);print(fs,r['load'],r['status'],flush=True)
  np.savez_compressed(ROOT/'results'/('bisector_'+str(fs)+'.npz'),xy=new.xy,tri=new.tri,material_id=new.material_id,x=x)
 write(ROOT/'results/bisector_plan.json',out)
if __name__=='__main__':main()
