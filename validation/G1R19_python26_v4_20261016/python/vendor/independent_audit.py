import math,numpy as np
from model_support.schema import topology
from model_support.p2 import gradients,bernstein
def strength(m,e,fs):
 c,p,d,_=m.materials[m.material_id[e]];p=math.atan(math.tan(math.radians(p))/fs);d=min(p,max(0.,math.radians(d)))
 if m.provenance.get('psi_policy')=='equal_reduced_phi':d=p
 eta=math.cos(d)*math.cos(p)/(1-math.sin(d)*math.sin(p))
 return eta*c/(fs*m.stress_scale),math.atan(eta*math.tan(p))

def zero_phi_vertex_weights(q):
 """Reconstruct triangle quadrature independently from its subtriangles."""
 counts={}
 for i in range(q):
  for j in range(q-i):
   cells=[[(i,j),(i+1,j),(i,j+1)]]
   if i+j<q-1:cells.append([(i+1,j),(i+1,j+1),(i,j+1)])
   for cell in cells:
    for node in cell:counts[node]=counts.get(node,0)+1
 return [(np.array([1-(i+j)/q,i/q,j/q]),counts[i,j]/(3*q*q))
         for i in range(q+1) for j in range(q+1-i)]

def audit(m,x,mode,fs,q,objective,load_id):
 area,grad,edges=topology(m);errors={};diss=0.;work=np.zeros(2)
 def check(k,v):errors[k]=max(errors.get(k,0.),float(v))
 if mode=='upper':
  nv=12*len(m.tri);v=x[:nv].reshape(len(m.tri),6,2);aux=nv;epigraph_diss=0.
  for e in range(len(m.tri)):
   c,p=strength(m,e,fs)
   for k in range(3):
    dv=v[e].T@gradients(np.eye(3)[k],grad[e]);tr=np.trace(dv);dev=dv[0,0]-dv[1,1];shear=dv[0,1]+dv[1,0]
    if p==0:check('flow',abs(tr))
    else:
     check('flow',math.sin(p)*math.hypot(dev,shear)-tr);term=area[e]*c*tr/(3*math.tan(p));diss+=term;epigraph_diss+=term
   if p==0:
    for bary,weight in zero_phi_vertex_weights(q):
     dv=v[e].T@gradients(bary,grad[e]);norm=math.hypot(dv[0,0]-dv[1,1],dv[0,1]+dv[1,0])
     rho=x[aux];aux+=1;check('epigraph',norm-rho)
     diss+=area[e]*c*weight*norm;epigraph_diss+=area[e]*c*weight*rho
   work+=area[e]/(3*m.stress_scale)*np.einsum('kj,jp->p',v[e,3:6],m.body[e])
  for ed in edges:
   e,f=ed.elements
   if f>=0:
    jump=v[f,ed.loc[1]]-v[e,ed.loc[0]];jt=jump@ed.tangent;jn=jump@ed.normal;c,p=strength(m,e,fs)
    if m.material_id[e]!=m.material_id[f]:
     check('bonded_interface',np.max(np.abs(jump)));continue
    if p==0:
     check('jump_normal',np.max(abs(jn)))
     for s in range(q):
      w=bernstein(s/q,(s+1)/q)
      for t in w@jt:
       rho=x[aux];aux+=1;check('epigraph',abs(t)-rho)
       diss+=ed.length*c/(3*q)*abs(t);epigraph_diss+=ed.length*c/(3*q)*rho
    else:
     for s in range(q):
      w=bernstein(s/q,(s+1)/q)
      for t,n in zip(w@jt,w@jn):check('jump',math.tan(p)*abs(t)-n)
     term=ed.length*c/math.tan(p)*np.array([1,1,4])@jn/6;diss+=term;epigraph_diss+=term
   else:
    vv=v[e,ed.loc[0]];work+=ed.length/(6*m.stress_scale)*np.einsum('k,kj,jp->p',np.array([1,1,4]),vv,ed.loads)
    for k in range(3):
     for j in range(2):
      if ed.kind=='fixed'or(ed.kind=='roller_x'and j==0)or(ed.kind=='roller_y'and j==1):check('boundary',abs(vv[k,j]))
  assert aux==len(x),'UPPER_AUXILIARY_LAYOUT'
  check('work',abs(work[1]-1));check('objective',abs(epigraph_diss-work[0]-objective))
  result={'pass':all(v<=(2e-6 if k=='objective'else 1e-7)for k,v in errors.items()),'errors':errors,'dissipation':float(diss),'work':work.tolist(),'value':float(diss-work[0])}
  if aux>nv:
   result.update(epigraph_value=float(epigraph_diss-work[0]),
                 dissipation_certificate='convex_norm_vertex_bound_and_Bernstein_jump_bound',
                 quadrature=q,auxiliary_variables=aux-nv)
  return result
 sig=x[:18*len(m.tri)].reshape(-1,6,3);load=x[load_id]
 for e in range(len(m.tri)):
  c,p=strength(m,e,fs)
  for sx,sy,t in sig[e]:check('yield',np.hypot(sx-sy,2*t)+math.sin(p)*(sx+sy)-2*c*math.cos(p))
  for bary in np.eye(3):
   g=np.zeros((6,2))
   for k in range(3):
    j=(k+1)%3;g[k]=2*bary[k]*grad[e,k];g[k+3]=2*(bary[k]*grad[e,j]+bary[j]*grad[e,k])
   for j in range(2):check('equilibrium',abs((m.body[e,j,0]+load*m.body[e,j,1])/m.stress_scale+sig[e,:,j]@g[:,j]+sig[e,:,2]@g[:,1-j]))
 for ed in edges:
  e,f=ed.elements
  for k in range(3):
   for j in range(2):
    value=sig[e,ed.loc[0][k],j]*ed.normal[j]+sig[e,ed.loc[0][k],2]*ed.normal[1-j]
    if f>=0:value-=sig[f,ed.loc[1][k],j]*ed.normal[j]+sig[f,ed.loc[1][k],2]*ed.normal[1-j];check('internal_traction',abs(value))
    elif ed.kind=='load'or(ed.kind=='roller_x'and j==1)or(ed.kind=='roller_y'and j==0):check('boundary_traction',abs(value-(ed.loads[j,0]+load*ed.loads[j,1])/m.stress_scale))
 return {'pass':max(errors.values())<=1e-7,'errors':errors,'load':float(load),'value':float(load)}
