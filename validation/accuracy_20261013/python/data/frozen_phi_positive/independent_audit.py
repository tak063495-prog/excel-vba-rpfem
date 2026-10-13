import math,numpy as np
from model_support.schema import topology
from model_support.p2 import gradients,bernstein
def strength(m,e,fs):
 c,p,d,_=m.materials[m.material_id[e]];p=math.atan(math.tan(math.radians(p))/fs);d=min(p,max(0.,math.radians(d)))
 if m.provenance.get('psi_policy')=='equal_reduced_phi':d=p
 eta=math.cos(d)*math.cos(p)/(1-math.sin(d)*math.sin(p))
 return eta*c/(fs*m.stress_scale),math.atan(eta*math.tan(p))
def audit(m,x,mode,fs,q,objective,load_id):
 area,grad,edges=topology(m);errors={};diss=0.;work=np.zeros(2)
 def check(k,v):errors[k]=max(errors.get(k,0.),float(v))
 if mode=='upper':
  v=x.reshape(len(m.tri),6,2)
  for e in range(len(m.tri)):
   c,p=strength(m,e,fs)
   for k in range(3):
    dv=v[e].T@gradients(np.eye(3)[k],grad[e]);tr=np.trace(dv);dev=dv[0,0]-dv[1,1];shear=dv[0,1]+dv[1,0]
    check('flow',math.sin(p)*math.hypot(dev,shear)-tr);diss+=area[e]*c*tr/(3*math.tan(p))
   work+=area[e]/(3*m.stress_scale)*np.einsum('kj,jp->p',v[e,3:6],m.body[e])
  for ed in edges:
   e,f=ed.elements
   if f>=0:
    jump=v[f,ed.loc[1]]-v[e,ed.loc[0]];jt=jump@ed.tangent;jn=jump@ed.normal;c,p=strength(m,e,fs)
    if m.material_id[e]!=m.material_id[f]:
     check('bonded_interface',np.max(np.abs(jump)));continue
    for s in range(q):
     w=bernstein(s/q,(s+1)/q)
     for t,n in zip(w@jt,w@jn):check('jump',math.tan(p)*abs(t)-n)
    diss+=ed.length*c/math.tan(p)*np.array([1,1,4])@jn/6
   else:
    vv=v[e,ed.loc[0]];work+=ed.length/(6*m.stress_scale)*np.einsum('k,kj,jp->p',np.array([1,1,4]),vv,ed.loads)
    for k in range(3):
     for j in range(2):
      if ed.kind=='fixed'or(ed.kind=='roller_x'and j==0)or(ed.kind=='roller_y'and j==1):check('boundary',abs(vv[k,j]))
  check('work',abs(work[1]-1));check('objective',abs(diss-work[0]-objective))
  return {'pass':all(v<=(2e-6 if k=='objective'else 1e-7)for k,v in errors.items()),'errors':errors,'dissipation':float(diss),'work':work.tolist(),'value':float(diss-work[0])}
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
