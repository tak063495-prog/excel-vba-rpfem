"""Python emulation of the VBA G1R20 Fix8 thin-layer chains (RPX_Adapt.Fix8ThinLayerChains).

Same structure/parameters as the VBA code: separate upper/lower lineages,
interleaved rounds (upper rounds 1..U, lower rounds at LOWER_START..),
upper = Doerfler(theta_u) directional one-edge split (+ mid-line on step 1),
cap check, flow-metric flips (5 passes) driven by the *prolonged* lineage
velocity (no extra solve, as in VBA); lower = mixed indicator (best-upper
kinematic density located by centroid + lower stress score, background 0),
Doerfler(theta_l) red split (+ mid-line on step 1), cap check, angular fan.
Each candidate gets its own q8 root (rel width <= FS_TOL 1e-4) and is
reported only from fields that pass the full original audit (plib).
Acceptance: valid root -> lineage advances; better -> new best witness.
Chain stops after 2 consecutive non-improving steps, on a failed build
(cap) or a failed root. Optional VBA budget emulation: a step whose upper
mesh exceeds --upper-fail-elements is treated as a failed solve (r=0).

usage: vba_chain.py CASE START OUT [--upper-fail-elements N]
  START = vba (Fix4 VBA witness meshes from bench26) | coarse (245-el input mesh)
"""
import sys,os,json,time,math,argparse
HERE=os.path.dirname(os.path.abspath(__file__))
PROTO=os.path.join(HERE,'..','..','G1R20_P06_meshproto_20261018','proto')
sys.path.insert(0,PROTO)
import numpy as np
import plib
from plib import core
import meshops as mo, mesh_geometry as mg, mesh_accuracy as ma
ap=argparse.ArgumentParser()
ap.add_argument('case');ap.add_argument('start');ap.add_argument('out')
ap.add_argument('--upper-fail-elements',type=int,default=0)
ap.add_argument('--upper-rounds',type=int,default=10);ap.add_argument('--lower-rounds',type=int,default=3)
ap.add_argument('--lower-start',type=int,default=2)
ap.add_argument('--upper-theta',type=float,default=.8);ap.add_argument('--lower-theta',type=float,default=.9)
ap.add_argument('--upper-cap',type=int,default=3500);ap.add_argument('--lower-cap',type=int,default=4000)
ap.add_argument('--mid-separate',action='store_true');ap.add_argument('--mid-every',action='store_true');ap.add_argument('--lower-own-upper',action='store_true');ap.add_argument('--flip-fresh',action='store_true')
ap.add_argument('--mid-ratio',type=float,default=.25);ap.add_argument('--target',type=float,default=.02)
a=ap.parse_args()
os.makedirs(a.out,exist_ok=True)
LOG=open(os.path.join(a.out,'log.jsonl'),'a')
T0=time.perf_counter()
def log(**k):
    k['t']=round(time.perf_counter()-T0,1);LOG.write(json.dumps(k)+'\n');LOG.flush();print(json.dumps({x:y for x,y in k.items() if x!='trials'}),flush=True)
m0,rec=plib.load(a.case)
BENCH='/workspace/excel-vba-rpfem/validation/G1R20_Fix4_bench26_20261017/cases'
def read_witness(side):
    raw=open(f'{BENCH}/{a.case}/witness_{side}.bin','rb').read();pos=[8]
    def take(dt,c):
        x=np.frombuffer(raw,dtype=dt,count=c,offset=pos[0]).copy();pos[0]+=x.nbytes;return x
    nn,ne,nm,nr,q=map(int,take('<i4',5));scale,factor,fs,obj,na=take('<f8',5)
    xy=take('<f8',nn*2).reshape(nn,2);tri=take('<i4',ne*5).reshape(ne,5)
    return core.mesh_from_arrays(xy,tri[:,:3],tri[:,3],rec),float(fs)
def area(m):
    p=m.xy[m.tri];u=p[:,1]-p[:,0];v=p[:,2]-p[:,0];return (u[:,0]*v[:,1]-u[:,1]*v[:,0])/2
def thin_mask(m):
    _,thick=mo.thin_materials(m,a.mid_ratio);return [i for i,t in enumerate(thick) if t<a.mid_ratio],thick
def solve_root(m,side,fs0):
    r=plib.root(m,side,fs0,rel_tol=1e-4,max_solves=20)
    tr=r['trials'];n=len(m.tri)
    info={'elements':n,'fs':r['fs'],'bad':r['bad'],'solves':r['solves'],'solve_s':round(r['solve_s'],2),
          'max_solve_s':round(r['max_solve_s'],2),'vars':tr[-1]['vars'],'cones':tr[-1]['cones'],
          'trials':[{k:t[k] for k in ('fs','value','witness','physical','status','iterations','solve_s','retry')} for t in tr]}
    ok=r['fs'] is not None and r['bad'] is not None and abs(r['bad']-r['fs'])<=1e-4*max(1,r['fs'])
    return ok,r['fs'],r['field'],info
# ---- P2 prolongation (exact restriction of the parent P2 field to nested children)
def p2_eval(vel_e,lam):
    l0,l1,l2=lam;N=np.array([l0*(2*l0-1),l1*(2*l1-1),l2*(2*l2-1),4*l0*l1,4*l1*l2,4*l2*l0]);return N@vel_e
def prolong(pm,pvel,cm):
    from matplotlib.tri import Triangulation
    T=Triangulation(pm.xy[:,0],pm.xy[:,1],pm.tri);finder=T.get_trifinder()
    pv=np.asarray(pvel)[:12*len(pm.tri)].reshape(-1,6,2);out=np.zeros((len(cm.tri),6,2))
    for e,t in enumerate(cm.tri):
        P=cm.xy[t];c=P.mean(0);pe=int(finder(c[0],c[1]))
        if pe<0:raise RuntimeError('prolong locate failed')
        A=pm.xy[pm.tri[pe]];M=np.array([[A[0,0],A[1,0],A[2,0]],[A[0,1],A[1,1],A[2,1]],[1,1,1]])
        pts=[P[0],P[1],P[2],(P[0]+P[1])/2,(P[1]+P[2])/2,(P[2]+P[0])/2]
        for j,q in enumerate(pts):out[e,j]=p2_eval(pv[pe],np.linalg.solve(M,[q[0],q[1],1.]))
    return out.reshape(-1)
def locate_density(um,uvel,lm):
    from matplotlib.tri import Triangulation
    su=core.indicator(um,uvel);dens=su/np.maximum(area(um),1e-300)
    T=Triangulation(um.xy[:,0],um.xy[:,1],um.tri);finder=T.get_trifinder()
    c=lm.xy[lm.tri].mean(1);idx=finder(c[:,0],c[:,1]);la=area(lm)
    s=np.where(idx>=0,dens[np.maximum(idx,0)]*la,0.);return s,int((idx<0).sum())
# ---- start
if a.start=='vba':
    um,ufs=read_witness('upper');lm,lfs=read_witness('lower')
else:
    um,ufs=m0,{'P06N':.88,'P06A':1.03}[a.case];lm,lfs=m0,{'P06N':.56,'P06A':.73}[a.case]
mask,thick=thin_mask(um)
if a.mid_separate:
    um,_=mo.split_edges(um,rec,mo.transverse_edges(um,materials=mask));lm,_=mo.split_edges(lm,rec,mo.transverse_edges(lm,materials=mask))
log(event='begin',case=a.case,start=a.start,args=vars(a),thickness=thick,mid_materials=mask,upper_elements=len(um.tri),lower_elements=len(lm.tri),vba_upper=ufs,vba_lower=lfs)
ok,U,ux,info=solve_root(um,'upper',ufs);log(event='start_root',side='upper',ok=ok,**info)
ok2,L,lx,info2=solve_root(lm,'lower',lfs);log(event='start_root',side='lower',ok=ok2,**info2)
best={'upper':(U,um,ux),'lower':(L,lm,lx)}
line={'upper':(U,um,ux),'lower':(L,lm,lx)}
alive={'upper':a.upper_rounds>0,'lower':a.lower_rounds>0};mid={'upper':not a.mid_separate,'lower':not a.mid_separate};stalls={'upper':0,'lower':0}
def gap():return plib.gap(best['lower'][0],best['upper'][0])
def step(side,rnd):
    fs,m,x=line[side];t=time.perf_counter();detail={}
    if side=='upper':
        score=core.indicator(m,x);E=mo.directional_edges(m,score,x,a.upper_theta);detail['directional_edges']=len(E)
        if mid['upper']:
            T=mo.transverse_edges(m,materials=mask);detail['mid_edges']=len(T-E);E=E|T
        cm,_=mo.split_edges(m,rec,E)
        if len(cm.tri)>a.upper_cap:return 0,dict(detail,reason='cap',elements=len(cm.tri))
        cv=(plib.evaluate(cm,'upper',fs)[1] if a.flip_fresh else None);cv=cv if cv is not None else prolong(m,x,cm);fm,finfo=mg.flow_metric_flip(cm,rec,cv,passes=5)
        detail['flips']=finfo.get('flips',finfo.get('changes')) if isinstance(finfo,dict) else None
        cm=fm
        if a.upper_fail_elements and len(cm.tri)>a.upper_fail_elements:
            return 0,dict(detail,reason='vba_budget_emulated',elements=len(cm.tri))
    else:
        if a.lower_own_upper:
            xu=plib.evaluate(m,'upper',fs*1.08)[1];su=np.asarray(core.indicator(m,xu),float);missed=0
        else:
            bu,bm,bx=best['upper'];su,missed=locate_density(bm,bx,m)
        sl=ma.lower_stress_score(m,x,fs,proximity_background=0.)
        mix=(su/su.sum() if su.sum()>0 else 0)+(sl/sl.sum() if sl.sum()>0 else 0)
        E=mo.red_edges(m,mo.dorfler(np.asarray(mix,float),a.lower_theta));detail.update(red_edges=len(E),unlocated=missed)
        if mid['lower']:
            T=mo.transverse_edges(m,materials=mask);detail['mid_edges']=len(T-E);E=E|T
        cm,_=mo.split_edges(m,rec,E)
        if len(cm.tri)>a.lower_cap:return 0,dict(detail,reason='cap',elements=len(cm.tri))
        cm,_=ma.angular_fan_refine(cm,rec)
    detail['build_s']=round(time.perf_counter()-t,1)
    ok,f,fx,info=solve_root(cm,side,fs)
    if not ok:return 0,dict(detail,reason='root_failed',**info)
    if side=='lower' and f>best['upper'][0]:return 0,dict(detail,reason='crossed',**info)
    if side=='upper' and f<best['lower'][0]:return 0,dict(detail,reason='crossed',**info)
    line[side]=(f,cm,fx);res=2
    if (side=='lower' and f>=best['lower'][0]) or (side=='upper' and f<=best['upper'][0]):best[side]=(f,cm,fx);res=1
    return res,dict(detail,**info)
U_R,L_R,LS=a.upper_rounds,a.lower_rounds,a.lower_start
for rnd in range(1,U_R+LS+L_R+1):
    if gap()<=a.target or not(alive['upper'] or alive['lower']):break
    for side in ('upper','lower'):
        if side=='upper':
            if not(alive['upper'] and rnd<=U_R):
                if rnd>U_R:alive['upper']=False
                continue
            k=rnd
        else:
            if not(alive['lower'] and LS<=rnd<LS+L_R):
                if rnd>=LS+L_R:alive['lower']=False
                continue
            k=rnd-LS+1
        before=best[side][0];r,detail=step(side,k)
        if r==0:alive[side]=False
        else:
            if not a.mid_every:mid[side]=False
            better=(best[side][0]<before*(1-1e-6)) if side=='upper' else (best[side][0]>before*(1+1e-6))
            stalls[side]=0 if better else stalls[side]+1
            if stalls[side]>=2:alive[side]=False
        log(event='step',side=side,round=k,result=r,best_lower=best['lower'][0],best_upper=best['upper'][0],gap=gap(),**detail)
        if gap()<=a.target:break
log(event='end',lower=best['lower'][0],upper=best['upper'][0],gap=gap(),lower_elements=len(best['lower'][1].tri),upper_elements=len(best['upper'][1].tri))
for s in ('lower','upper'):
    from accuracy_driver import save_mesh
    save_mesh(os.path.join(a.out,f'best_{s}.npz'),best[s][1])
