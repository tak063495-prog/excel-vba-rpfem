import os
for k in ['OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS']:os.environ[k]='1'
from pathlib import Path
import sys,json,time,copy
import numpy as np
from scipy import sparse
import clarabel
R=Path(__file__).resolve().parent
P=R/'validation_runtime'
if not P.exists():P=Path('C:/Users/link_/Desktop/RPFEM_G1R9_Speedup_20261006/validation_runtime')
sys.path.insert(0,str(P/'python'))
from common import Mesh,CASES,model_key,save_mesh,physical_audit,topology,assemble_lower
from conic_lp import program,affine_reduce,cone_error
from audit_precision import audit_precision
from inspect_book import cells,z
(R/'results').mkdir(exist_ok=True)
def cross2(a,b):return a[...,0]*b[...,1]-a[...,1]*b[...,0]

def saved_mesh():
 nd=cells('xl/worksheets/sheet7.xml');ed=cells('xl/worksheets/sheet8.xml')
 rn=sum(k.startswith('F') and k[1:].isdigit() and int(k[1:])>=2 for k in nd)
 ne=sum(k.startswith('AA') and k[2:].isdigit() and int(k[2:])>=2 for k in ed)
 xy=np.array([[float(nd[f'G{i}']),float(nd[f'H{i}'])] for i in range(2,rn+2)])
 tri=np.array([[int(ed[f'{c}{i}'])-1 for c in ['AB','AC','AD']] for i in range(2,ne+2)])
 mat=np.array([int(ed[f'AE{i}'])-1 for i in range(2,ne+2)])
 counts={}
 for t in tri:
  for a,b in zip(t,np.roll(t,-1)):
   key=tuple(sorted((int(a),int(b))));counts[key]=counts.get(key,0)+1
 boundary={}
 for key,count in counts.items():
  if count==1:
   pts=xy[list(key)]
   kind='fixed' if np.max(abs(pts[:,1]))<1e-10 else 'roller_x' if np.max(abs(pts[:,0]))<1e-10 or np.max(abs(pts[:,0]-28))<1e-10 else 'load'
   boundary[key]={'kind':kind,'rigid':-1,'loads':np.zeros((2,2))}
 mats=np.array([[20.,35.,0.,19.],[0.,25.,0.,19.],[10.,35.,0.,19.]])
 body=np.zeros((ne,2,2));body[:,1,1]=-19.
 m=Mesh(xy,tri,mats,mat,np.full(ne,-1),body,boundary,20.,model_key(CASES['P06N']),provenance={'psi_policy':'hold_initial','case':'P06N','source':'saved_G1R9_lower_geometry'})
 return m.validate()

def refine(m,selected,red=True):
 xy=m.xy.tolist();marks={};boundary=copy.deepcopy(m.boundary)
 for e in sorted(selected):
  t=m.tri[e]
  for a,b in zip(t,np.roll(t,-1)):
   a,b=int(a),int(b);key=tuple(sorted((a,b)))
   if key not in marks:
    j=len(xy);xy.append(((m.xy[a]+m.xy[b])/2).tolist());marks[key]=j
    if key in boundary:
     rec=boundary.pop(key);boundary[tuple(sorted((a,j)))]=copy.deepcopy(rec);boundary[tuple(sorted((j,b)))]=copy.deepcopy(rec)
 tris=[];mid=[];bodies=[]
 for e,t0 in enumerate(m.tri):
  t=list(map(int,t0));ids=[marks.get(tuple(sorted((t[k],t[(k+1)%3]))),-1) for k in range(3)];hit=[k for k,v in enumerate(ids) if v>=0]
  if not hit:pieces=[t]
  elif not red:
   j=len(xy);xy.append(np.mean(m.xy[t],axis=0).tolist());poly=[]
   for k in range(3):
    poly.append(t[k])
    if ids[k]>=0:poly.append(ids[k])
   pieces=[[poly[k],poly[(k+1)%len(poly)],j] for k in range(len(poly))]
  elif len(hit)==3:
   a,b,c=t;ab,bc,ca=ids;pieces=[[a,ab,ca],[ab,b,bc],[ca,bc,c],[ab,bc,ca]]
  elif len(hit)==1:
   k=hit[0];a,b,c=t[k],t[(k+1)%3],t[(k+2)%3];ab=ids[k];pieces=[[a,ab,c],[ab,b,c]]
  else:
   # Rotate so marked edges are a-b and c-a, around vertex a.
   k=next(k for k in range(3) if ids[k]>=0 and ids[(k+2)%3]>=0)
   a,b,c=t[k],t[(k+1)%3],t[(k+2)%3];ab,ca=ids[k],ids[(k+2)%3]
   if np.sum((np.array(xy[ab])-np.array(xy[c]))**2)<=np.sum((np.array(xy[ca])-np.array(xy[b]))**2):pieces=[[a,ab,ca],[ab,b,c],[ab,c,ca]]
   else:pieces=[[a,ab,ca],[ab,b,ca],[b,c,ca]]
  for piece in pieces:
   p=np.array([xy[i] for i in piece]);cross=cross2(p[1]-p[0],p[2]-p[0])
   if cross<0:piece[1],piece[2]=piece[2],piece[1]
   assert abs(cross)>1e-14
   tris.append(piece);mid.append(m.material_id[e]);bodies.append(m.body[e].copy())
 out=Mesh(np.array(xy),np.array(tris),m.materials.copy(),np.array(mid),np.full(len(tris),-1),np.array(bodies),boundary,m.stress_scale,m.physical_model_id,provenance=m.provenance.copy()).validate()
 assert np.max(abs(material_area(m)-material_area(out)))<1e-9
 return out

def material_area(m):
 p=m.xy[m.tri];ar=np.abs(cross2(p[:,1]-p[:,0],p[:,2]-p[:,0]))/2
 return np.bincount(m.material_id,weights=ar,minlength=3)

def probe(m,label,fs,mode,q=4):
 original=program(m,fs,mode,q);E,b,G,h,dims,c=original
 er,br,gr,hr,dr,T,off,trace=affine_reduce(E,b,G,h,dims)
 A=sparse.vstack([er,-gr],format='csc');rhs=np.r_[br,hr];cost=np.asarray(c@T).ravel()
 cones=[clarabel.ZeroConeT(er.shape[0])]+[clarabel.SecondOrderConeT(int(d)) if d>1 else clarabel.NonnegativeConeT(1) for d in dr]
 st=clarabel.DefaultSettings();st.verbose=False;st.max_iter=250;st.time_limit=45.;st.tol_feas=1e-9;st.tol_gap_abs=1e-9;st.tol_gap_rel=1e-9
 t=time.perf_counter();sol=clarabel.DefaultSolver(sparse.csc_matrix((A.shape[1],A.shape[1])),cost,A,rhs,cones,st).solve();x=np.asarray(T@np.asarray(sol.x)+off).ravel()
 au=physical_audit(m,x,mode,fs,q);eq=float(np.max(abs(E@x-b)));co=cone_error(G@x+h,dims)
 meaningful=abs(x[-1]-1)<1e-9 if mode=='lower' else au['value']<=1-1e-7 and float(c@x)<=1-1e-7
 rawmax=max(au.get('errors',{}).values(),default=1.)*(m.stress_scale if mode=='lower' else 1)
 accepted=au['pass'] and eq<=1e-7 and co<=1e-7 and rawmax<=1e-7 and meaningful
 rec={'mesh':label,'elements':len(m.tri),'fs':fs,'mode':mode,'q':q,'status':str(sol.status),'seconds':time.perf_counter()-t,'accepted_witness':bool(accepted),'value':float(c@x) if mode=='upper' else float(x[-1]),'physical_audit':au,'eq':eq,'cone':co,'rawmax':float(rawmax),'scope':'AUDITED_SIDE_WITNESS_NOT_EXACT_OPTIMUM_OR_ROOT'}
 if accepted:
  save_mesh(R/'results'/f'{label}_{mode}_{fs}_{q}.npz',m,x=x,fs=fs,mode=mode,q=q)
  if mode=='lower':
   precise=audit_precision(m,x,fs,mode);rec['precision']=precise;assert precise['numerical_audit_pass']
 print(json.dumps({k:v for k,v in rec.items() if k not in ['physical_audit','precision']},ensure_ascii=False),flush=True)
 return rec

if __name__=='__main__':
 base=saved_mesh();weak=set(np.flatnonzero(base.materials[base.material_id,0]==0));variants={'base781':base,'weak_red':refine(base,weak,True),'weak_fan':refine(base,weak,False)}
 records=[]
 for label,m in variants.items():
  save_mesh(R/'results'/(label+'_mesh.npz'),m)
  print('MESH',label,len(m.tri),len(m.xy),material_area(m),flush=True)
  for fs,mode in [(.68,'lower'),(.8,'upper')]:
   records.append(probe(m,label,fs,mode))
   (R/'results/experiments.json').write_text(json.dumps(records,ensure_ascii=False,indent=2),encoding='utf-8')
