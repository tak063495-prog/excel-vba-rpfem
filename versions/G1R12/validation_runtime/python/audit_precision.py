"""Independent 50-digit audit of stored P06 P2 fields.
No assembled E/G or solver residuals are used. Stored binary64 inputs are lifted
exactly to Decimal. Material trig is evaluated independently with mpmath (60 dps).
This is a numerical audit, not an interval-arithmetic proof.
"""
from decimal import Decimal,localcontext
import mpmath as mp
import numpy as np
from common import *

D=lambda x:Decimal(float(x))
Z=Decimal(0);O=Decimal(1)

def audit_precision(m,x,fs,mode):
 with localcontext() as ctx:
  ctx.prec=50
  return _audit(m,x,fs,mode)

def _audit(m,x,fs,mode):
 mp.mp.dps=65;xy=[[D(a) for a in row]for row in m.xy];scale=D(m.stress_scale);F=D(fs)
 pars=[]
 for cc,pp,ps,g in m.materials:
  p=mp.atan(mp.tan(mp.mpf(float(pp))*mp.pi/180)/mp.mpf(float(fs)))
  d=p if m.provenance['psi_policy']=='equal_reduced_phi' else min(p,max(mp.mpf(0),mp.mpf(float(ps))*mp.pi/180))
  eta=mp.cos(p)*mp.cos(d)/(1-mp.sin(p)*mp.sin(d));pq=mp.atan(eta*mp.tan(p))
  pars.append(tuple(Decimal(mp.nstr(a,62)) for a in [eta*mp.mpf(float(cc))/mp.mpf(float(fs)),mp.sin(pq),mp.cos(pq),mp.tan(pq)]))
 grads=[];areas=[];edges={}
 for e,t in enumerate(m.tri):
  pt=[xy[i]for i in t];det=(pt[1][0]-pt[0][0])*(pt[2][1]-pt[0][1])-(pt[1][1]-pt[0][1])*(pt[2][0]-pt[0][0]);assert det>0
  areas.append(det/2);grads.append([[(pt[(k+1)%3][1]-pt[(k+2)%3][1])/det,(pt[(k+2)%3][0]-pt[(k+1)%3][0])/det]for k in range(3)])
  for k in range(3):
   a,b=int(t[k]),int(t[(k+1)%3]);key=tuple(sorted((a,b)));loc=[k,(k+1)%3,3+k]if a<b else[(k+1)%3,k,3+k]
   edges.setdefault(key,[]).append((e,loc));assert len(edges[key])<=2
 err={};where={}
 def check(k,v,loc):
  v=max(Z,v)
  if v>err.get(k,Decimal(-1)):err[k]=v;where[k]=loc
 vec=np.asarray(x,float)
 if not np.isfinite(vec).all():raise ValueError('NONFINITE_FIELD')
 if mode=='lower':
  sig=[[[D(v)*scale for v in abc]for abc in row]for row in vec[:18*len(m.tri)].reshape(-1,6,3)]
  lam=D(vec[18*len(m.tri)]);check('load_one',abs(lam-1),'load')
  for e in range(len(m.tri)):
   cc,sp,cp,tp=pars[m.material_id[e]]
   for k,(sx,sy,tau) in enumerate(sig[e]):check('yield',((sx-sy)**2+4*tau*tau).sqrt()+sp*(sx+sy)-2*cc*cp,[e,k])
   for r in range(3):
    dg=[[Z,Z]for i in range(6)]
    for k in range(3):
     j=(k+1)%3
     for d in range(2):
      dg[k][d]=2*(O if k==r else Z)*grads[e][k][d]
      dg[k+3][d]=2*((O if k==r else Z)*grads[e][j][d]+(O if j==r else Z)*grads[e][k][d])
    for d in range(2):
     v=D(m.body[e,d,0])+lam*D(m.body[e,d,1])+sum((sig[e][i][d]*dg[i][d]+sig[e][i][2]*dg[i][1-d] for i in range(6)),Z)
     check('equilibrium',abs(v),[e,r,d])
 else:
  vel=[[[D(v)for v in ab]for ab in row]for row in vec[:12*len(m.tri)].reshape(-1,6,2)];diss=Z;work=[Z,Z];abswork=[Z,Z]
  for e in range(len(m.tri)):
   cc,sp,cp,tp=pars[m.material_id[e]]
   for r in range(3):
    dg=[]
    for k in range(3):dg.append([(4*(O if k==r else Z)-1)*a for a in grads[e][k]])
    for k in range(3):
     j=(k+1)%3;dg.append([4*((O if k==r else Z)*grads[e][j][d]+(O if j==r else Z)*grads[e][k][d])for d in range(2)])
    dv=[[sum((vel[e][i][a]*dg[i][b]for i in range(6)),Z)for b in range(2)]for a in range(2)]
    tr=dv[0][0]+dv[1][1];dev=dv[0][0]-dv[1][1];sh=dv[0][1]+dv[1][0]
    check('flow',sp*(dev*dev+sh*sh).sqrt()-tr,[e,r]);diss+=areas[e]*(cc/scale)*tr/(3*tp)
   for ph in range(2):
    w=areas[e]/(3*scale)*sum((vel[e][i][j]*D(m.body[e,j,ph])for i in [3,4,5]for j in range(2)),Z);work[ph]+=w;abswork[ph]+=abs(w)
 for key,sides in edges.items():
  a,b=key;delta=[xy[b][i]-xy[a][i]for i in range(2)];length=sum((v*v for v in delta),Z).sqrt();t=[v/length for v in delta];n=[t[1],-t[0]];e,loc=sides[0]
  center=[sum((xy[k][j]for k in m.tri[e]),Z)/3-(xy[a][j]+xy[b][j])/2 for j in range(2)]
  if sum((center[j]*n[j]for j in range(2)),Z)>0:n=[-a for a in n]
  f,locf=sides[1]if len(sides)==2 else(-1,None)
  if mode=='lower':
   for k in range(3):
    for j in range(2):
     val=sig[e][loc[k]][j]*n[j]+sig[e][loc[k]][2]*n[1-j]
     if f>=0:
      val-=sig[f][locf[k]][j]*n[j]+sig[f][locf[k]][2]*n[1-j];check('internal_traction',abs(val),[a,b,k,j])
     else:
      bd=m.boundary[key];kind=bd['kind']
      if kind=='load'or(kind=='roller_x'and j==1)or(kind=='roller_y'and j==0):
       check('boundary_traction',abs(val-D(bd['loads'][j,0])-lam*D(bd['loads'][j,1])),[a,b,k,j])
  elif f>=0:
   jump=[[vel[f][locf[k]][j]-vel[e][loc[k]][j]for j in range(2)]for k in range(3)]
   if m.material_id[e]!=m.material_id[f]:
    check('bonded_interface',max(abs(a)for row in jump for a in row),[a,b]);continue
   cc,sp,cp,tp=pars[m.material_id[e]];jn=[sum((row[j]*n[j]for j in range(2)),Z)for row in jump];jt=[sum((row[j]*t[j]for j in range(2)),Z)for row in jump]
   for sgn in [-1,1]:
    vv=[jn[k]+sgn*tp*jt[k]for k in range(3)];c=vv[0];aa=2*(vv[0]+vv[1]-2*vv[2]);bb=-3*vv[0]-vv[1]+4*vv[2]
    vals=[vv[0],vv[1]]
    if aa!=0:
     ss=-bb/(2*aa)
     if 0<ss<1:vals.append(aa*ss*ss+bb*ss+c)
    check('jump_whole_edge',-min(vals),[a,b,sgn])
   diss+=length*(cc/scale)*(jn[0]+jn[1]+4*jn[2])/(6*tp)
  else:
   bd=m.boundary[key];kind=bd['kind'];vv=[vel[e][k]for k in loc]
   for k in range(3):
    for j in range(2):
     if kind=='fixed'or(kind=='roller_x'and j==0)or(kind=='roller_y'and j==1):check('boundary_velocity',abs(vv[k][j]),[a,b,k,j])
   for ph in range(2):
    w=length/(6*scale)*sum((Decimal([1,1,4][k])*vv[k][j]*D(bd['loads'][j,ph])for k in range(3)for j in range(2)),Z);work[ph]+=w;abswork[ph]+=abs(w)
 if mode=='upper':check('reference_work',abs(work[1]-1),'work')
 rr={'case':m.provenance['case'],'mode':mode,'fs':fs,'elements':len(m.tri),'decimal_precision':50,'independent_trig_digits':65,'raw_errors':{k:float(v)for k,v in err.items()},'max_locations':where,'area':float(sum(areas,Z)),'numerical_audit_pass':all(v<=Decimal('1e-7')for v in err.values()),'meaning':'Numerical audit of stored binary64 field, not rigorous real interval proof.'}
 if mode=='upper':rr.update(dissipation=float(diss),fixed_work=float(work[0]),reference_work=float(work[1]),abs_work=float(abswork[1]),collapse_condition=bool(diss-work[0]<=1-Decimal('1e-7')))
 else:rr.update(load_factor=float(lam),stress_scale=float(scale))
 rr['accepted']=rr['numerical_audit_pass']and(rr.get('collapse_condition',True))
 return rr

def geometry_audit(m):
 from shapely.geometry import Polygon,LineString
 from shapely.ops import unary_union
 c=CASES[m.provenance['case']];polys=[Polygon(m.xy[t])for t in m.tri];union=unary_union(polys);outline=Polygon(c['outline']);errs={}
 for k,r in enumerate(c['regions']):
  pp=unary_union([p for p,i in zip(polys,m.material_id)if i==k]);errs[r['name']]=float(pp.symmetric_difference(Polygon(r['xy'])).area)
 ar=topology(m)[0];weight=-float(np.dot(ar,m.body[:,1,1]));cracks=[];kinds={};bad_bcs=[]
 for ed in topology(m)[2]:
  if ed.elements[1]<0:
   seg=LineString(m.xy[[ed.a,ed.b]]);kinds[ed.kind]=kinds.get(ed.kind,0.)+ed.length
   if seg.difference(outline.boundary.buffer(1e-8)).length>1e-8:cracks.append([ed.a,ed.b])
   a,b=m.xy[[ed.a,ed.b]]
   expected='fixed' if abs(a[1])+abs(b[1])<1e-8 else ('roller_x' if ((abs(a[0])+abs(b[0])<1e-8)or(abs(a[0]-28)+abs(b[0]-28)<1e-8)) else 'load')
   if ed.kind!=expected or np.any(ed.loads!=0):bad_bcs.append([ed.a,ed.b,ed.kind,expected])
 c_m=np.array([[r['c'],r['phi'],r['psi'],r['gamma']]for r in c['materials']]);matmatch=np.array_equal(c_m,m.materials)
 rec={'case':m.provenance['case'],'nodes':len(m.xy),'elements':len(m.tri),'area':float(ar.sum()),'region_symdiff':errs,'outline_symdiff':float(union.symmetric_difference(outline).area),'overlap_area':float(ar.sum()-union.area),'weight':weight,'internal_cracks':cracks,'incorrect_boundary_conditions':bad_bcs,'boundary_lengths':kinds,'materials_exact':matmatch,'body_exact':bool(np.all(m.body[:,0,:]==0)and np.all(m.body[:,1,0]==0)and np.all(m.body[:,1,1]==-19))}
 rec['pass']=bool(matmatch and rec['body_exact']and max(errs.values())<1e-8 and rec['outline_symdiff']<1e-8 and abs(rec['overlap_area'])<1e-8 and not cracks and not bad_bcs and abs(weight-5747.5)<1e-7)
 return rec
if __name__=='__main__':
 import sys
 for name in sys.argv[1:]:
  p=Path(name);z=np.load(p,allow_pickle=False);case='P06A'if 'P06A'in p.name else'P06N';m=load_mesh(case,path=p)
  r=audit_precision(m,z['x'],float(z['fs']),str(z['mode']));r['geometry']=geometry_audit(m);dump(ROOT/'results'/f'{p.stem}_50digit.json',r);print(p.name,r,flush=True)
