"""VBA RPX_Upper phi=0 epigraph/quadrature path, no artificial friction."""
import numpy as np
from scipy import sparse
from model_support.schema import topology
from model_support.p2 import gradients,bernstein
from model_support.upper_assembly import reduced_cohesion,Program
def weights(q):
 w=np.zeros((q+1,q+1))
 for i in range(q):
  for j in range(q-i):
   w[i,j]+=1;w[i+1,j]+=1;w[i,j+1]+=1
   if i+j<q-1:w[i+1,j]+=1;w[i+1,j+1]+=1;w[i,j+1]+=1
 return w
def assemble_zero(m,fs,q):
 assert np.all(m.materials[:,1]==0) and np.all(m.materials[:,2]==0)
 area,grad,edges=topology(m);nv=12*len(m.tri);n=nv;eq=[];rows=[];dims=[];cost={};work={};w=weights(q)
 def strain(e,bary):
  g=gradients(bary,grad[e]);ids=np.arange(12*e,12*e+12);dev=g.copy();dev[:,1]*=-1
  return dict(zip(ids,g.ravel())),dict(zip(ids,dev.ravel())),dict(zip(ids,g[:,::-1].ravel()))
 def var(price):
  nonlocal n
  i=n;n+=1;cost[i]=price;return i
 for e in range(len(m.tri)):
  c=reduced_cohesion(m,e,fs)
  for bary in np.eye(3):eq.append(strain(e,bary)[0])
  for i in range(q+1):
   for j in range(q+1-i):
    bary=np.array([1-(i+j)/q,i/q,j/q]);_,dev,shear=strain(e,bary);rho=var(area[e]*c*w[i,j]/(3*q*q));rows.extend([{rho:1.},dev,shear]);dims.append(3)
  for k in range(3,6):
   for j in range(2):work[12*e+2*k+j]=area[e]*m.body[e,j,1]/(3*m.stress_scale)
 for edge in edges:
  e,f=edge.elements
  if f>=0:
   jump=[{12*f+2*edge.loc[1][k]+j:1.,12*e+2*edge.loc[0][k]+j:-1.}for k in range(3)for j in range(2)]
   if m.material_id[e]!=m.material_id[f]:eq.extend(jump);continue
   normal=[];tangent=[]
   for k in range(3):
    normal.append({i:v*edge.normal[j]for j in range(2)for i,v in jump[2*k+j].items()})
    tangent.append({i:v*edge.tangent[j]for j in range(2)for i,v in jump[2*k+j].items()})
   eq.extend(normal);c=reduced_cohesion(m,e,fs)
   for s in range(q):
    for bw in bernstein(s/q,(s+1)/q):
     rho=var(edge.length*c/(3*q))
     for sign in [-1.,1.]:rows.append({rho:1.,**{i:sign*bw[k]*v for k in range(3)for i,v in tangent[k].items()}});dims.append(1)
  else:
   for k,wt in enumerate([1,1,4]):
    for j in range(2):
     i=12*e+2*edge.loc[0][k]+j;work[i]=work.get(i,0)+edge.length*wt*edge.loads[j,1]/(6*m.stress_scale)
     if edge.kind=='fixed'or(edge.kind=='roller_x'and j==0)or(edge.kind=='roller_y'and j==1):eq.append({i:1.})
 eq.append(work)
 def matrix(rr):
  ii=[];jj=[];vv=[]
  for r,row in enumerate(rr):
   for i,v in row.items():
    if v:ii.append(r);jj.append(i);vv.append(v)
  return sparse.coo_matrix((vv,(ii,jj)),shape=(len(rr),n)).tocsc()
 E=matrix(eq);norm=np.sqrt(np.asarray(E.power(2).sum(axis=1)).ravel());assert np.all(norm>0)
 objective=np.zeros(n)
 for i,v in cost.items():objective[i]=v
 rhs=np.zeros(len(eq));rhs[-1]=1.;workvec=np.zeros(n)
 for i,v in work.items():workvec[i]=v
 return Program(sparse.diags(1/norm)@E,rhs/norm,matrix(rows),dims,objective,workvec,norm)
def audit_zero(m,x,fs,q,objective):
 area,grad,edges=topology(m);v=x[:12*len(m.tri)].reshape(-1,6,2);err={};work=0.;diss=0.;w=weights(q)
 def check(k,value):err[k]=max(err.get(k,0.),float(value))
 for e in range(len(m.tri)):
  c=m.materials[m.material_id[e],0]/(fs*m.stress_scale)
  for bary in np.eye(3):check('flow',abs(np.trace(v[e].T@gradients(bary,grad[e]))))
  for i in range(q+1):
   for j in range(q+1-i):
    dv=v[e].T@gradients(np.array([1-(i+j)/q,i/q,j/q]),grad[e]);diss+=area[e]*c*w[i,j]/(3*q*q)*np.hypot(dv[0,0]-dv[1,1],dv[0,1]+dv[1,0])
  work+=area[e]/(3*m.stress_scale)*np.sum(v[e,3:6]@m.body[e,:,1])
 for edge in edges:
  e,f=edge.elements
  if f>=0:
   jump=v[f,edge.loc[1]]-v[e,edge.loc[0]]
   if m.material_id[e]!=m.material_id[f]:check('bonded_interface',np.max(abs(jump)));continue
   check('jump_normal',np.max(abs(jump@edge.normal)));c=m.materials[m.material_id[e],0]/(fs*m.stress_scale)
   for s in range(q):diss+=edge.length*c/(3*q)*np.sum(abs(bernstein(s/q,(s+1)/q)@(jump@edge.tangent)))
  else:
   vv=v[e,edge.loc[0]];work+=edge.length/(6*m.stress_scale)*np.array([1,1,4])@(vv@edge.loads[:,1])
   for k in range(3):
    for j in range(2):
     if edge.kind=='fixed'or(edge.kind=='roller_x'and j==0)or(edge.kind=='roller_y'and j==1):check('boundary',abs(vv[k,j]))
 check('work',abs(work-1));check('objective',abs(diss-objective))
 return {'pass':all(v<=(2e-6 if k=='objective'else 1e-7)for k,v in err.items()),'errors':err,'value':float(diss),'work':float(work),'quadrature':q}
