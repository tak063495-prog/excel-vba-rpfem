"""Test small topology changes before any VBA implementation."""
import argparse,json,copy,time,numpy as np
from support import ROOT,write
from diagnose import mesh_from_saved,mesh_from_arrays,solve
from model_support.schema import topology

def centroid(m,rec,elements):
 xy=m.xy.tolist();tri=[];ids=[]
 for e,t in enumerate(m.tri):
  if e not in elements:tri.append(t.tolist());ids.append(int(m.material_id[e]));continue
  mid=len(xy);xy.append(m.xy[t].mean(axis=0).tolist())
  for k in range(3):tri.append([int(t[k]),int(t[(k+1)%3]),mid]);ids.append(int(m.material_id[e]))
 return mesh_from_arrays(np.asarray(xy),np.asarray(tri,np.int32),np.asarray(ids,np.int32),rec)

def flip(m,rec,a,b):
 shared=set(m.tri[a])&set(m.tri[b]);assert len(shared)==2
 u=next(i for i in m.tri[a] if i not in shared);v=next(i for i in m.tri[b] if i not in shared);p,q=shared
 t=m.tri.copy();t[a]=[u,v,p];t[b]=[v,u,q]
 for e in [a,b]:
  pts=m.xy[t[e]]
  a=pts[1]-pts[0];b=pts[2]-pts[0]
  if a[0]*b[1]-a[1]*b[0]<0:t[e,[1,2]]=t[e,[2,1]]
 result=mesh_from_arrays(m.xy.copy(),t,m.material_id.copy(),rec)
 assert abs(topology(result)[0].sum()-topology(m)[0].sum())<1e-9
 return result

def main():
 p=argparse.ArgumentParser();p.add_argument('--tag',default='mesh_experiments');args=p.parse_args();folder=ROOT/'results'/args.tag;folder.mkdir(exist_ok=False)
 m,rec=mesh_from_saved();models={'centroid65':centroid(m,rec,{65}),'centroid51':centroid(m,rec,{51}),'centroid18_51_65':centroid(m,rec,{18,51,65}),'flip18_51':flip(m,rec,18,51)}
 out=[]
 for name,new in models.items():
  row={'name':name,'elements':len(new.tri),'mesh_id':new.mesh_id,'points':[]};out.append(row)
  a,_,_=topology(new);assert np.max(abs(np.bincount(new.material_id,weights=a,minlength=3)-np.array(rec['region_areas'])))<1e-9
  for fs in [1.,.5,.125]:
   r,x,pr,z=solve(new,fs,'lower');row['points'].append(r);np.savez_compressed(folder/(name+'_'+str(fs)+'.npz'),xy=new.xy,tri=new.tri,material_id=new.material_id,x=x)
   write(folder/'summary.json',{'experiments':out,'scope':'243 base plus2 or6 elements only; same physical inputs'})
   print(name,fs,r['status'],r['load'],r.get('physical_audit'),flush=True)
  r,x,pr,z=solve(new,1.,'upper');row['upper_at_1']=r;write(folder/'summary.json',{'experiments':out})
 print('LOCAL MESH EXPERIMENTS COMPLETE',flush=True)
if __name__=='__main__':main()
