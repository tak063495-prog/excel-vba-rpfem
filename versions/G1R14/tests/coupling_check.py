# Reference check for RPX_Water (one-way coupling, effective stress): mirrors RPX_WaterApply.
import sys, math, numpy as np
sys.path.insert(0, '..')
from seep_check import solve
from scipy.spatial import Delaunay
POLY=[(0,0),(20,0),(20,4),(16,4),(10,10),(0,10)]
GAMMA=20.0; GW=9.81
class _P:
    def contains_point(self,p):
        x,y=p; c=False
        for (ax,ay),(bx,by) in zip(POLY,POLY[1:]+POLY[:1]):
            if (ay>y)!=(by>y) and x<ax+(y-ay)*(bx-ax)/(by-ay): c=not c
        return c
poly=_P()
def on_boundary(x,y,tol=1e-9):
    for (ax,ay),(bx,by) in zip(POLY,POLY[1:]+POLY[:1]):
        L=math.hypot(bx-ax,by-ay); cr=abs((bx-ax)*(y-ay)-(by-ay)*(x-ax))/L
        t=((x-ax)*(bx-ax)+(y-ay)*(by-ay))/L**2
        if cr<tol and -tol<=t<=1+tol: return True
    return False
def mesh(h, shift=0.0):
    pts=[]
    for (ax,ay),(bx,by) in zip(POLY,POLY[1:]+POLY[:1]):
        n=max(1,int(math.ceil(math.hypot(bx-ax,by-ay)/h)))
        for i in range(n): pts.append((ax+(bx-ax)*i/n, ay+(by-ay)*i/n))
    y=h*0.5+shift
    j=0
    while y<10:
        x=h*(0.5*(j%2))+shift
        while x<20:
            if poly.contains_point((x,y)) and min(math.hypot(x-px,y-py) for px,py in pts)>0.4*h and not on_boundary(x,y,0.3*h): pts.append((x,y))
            x+=h
        y+=h*0.866; j+=1
    P=np.array(pts); T=Delaunay(P).simplices
    keep=[]
    for t in T:
        c=P[t].mean(0)
        if poly.contains_point(c):
            a,b,cc=t
            if (P[b,0]-P[a,0])*(P[cc,1]-P[a,1])-(P[cc,0]-P[a,0])*(P[b,1]-P[a,1])<0: t=[a,cc,b]
            keep.append(list(t))
    return P,keep
def seep(P,T,case):
    n=len(P); bt=[0]*n; bv=[0.0]*n
    for i,(x,y) in enumerate(P):
        if not on_boundary(x,y): continue
        if case=='full': bt[i]=1; bv[i]=12.0
        elif case=='still6':
            if y<=6+1e-9 and y>1e-9 and x>1e-9: bt[i]=1; bv[i]=6.0
            elif x>1e-9 and y>1e-9: bt[i]=2
            if abs(x)<1e-9 and y<=6: bt[i]=1; bv[i]=6.0
        elif case=='flow':
            if abs(x)<1e-9: bt[i]=1; bv[i]=10.0
            elif y>1e-9 and not (abs(y-10)<1e-9 and x<10):
                if y<=2+1e-9: bt[i]=1; bv[i]=2.0
                else: bt[i]=2
    h,q,it,kr=solve([tuple(p) for p in P],[tuple(t) for t in T],[0]*len(T),[1e-5],[1e-5],bt,bv)
    return np.array(h),it
def head_at(Ps,Ts,hs,x,y):   # mirrors RPX_WaterHeadAt (best = max of min barycentric, clamped)
    best=None;bm=-1e300
    for t in Ts:
        (x1,y1),(x2,y2),(x3,y3)=Ps[t]
        det=(x2-x1)*(y3-y1)-(x3-x1)*(y2-y1)
        l1=((x-x1)*(y3-y1)-(x3-x1)*(y-y1))/det; l2=((x2-x1)*(y-y1)-(x-x1)*(y2-y1))/det; l=[1-l1-l2,l1,l2]
        if min(l)>bm: bm=min(l); best=(t,l)
    t,l=best
    if bm<0: l=[max(v,0) for v in l]; s=sum(l); l=[v/s for v in l]
    return sum(l[k]*hs[t[k]] for k in range(3))
def apply(PB,TB,Ps,Ts,hs,suction=False):  # mirrors RPX_WaterApply (fixed gravity mode)
    u=np.array([GW*(head_at(Ps,Ts,hs,x,y)-y) for x,y in PB])
    if not suction: u=np.maximum(u,0)
    body=[];area=[]
    for t in TB:
        (x1,y1),(x2,y2),(x3,y3)=PB[t]; det=(x2-x1)*(y3-y1)-(x3-x1)*(y2-y1)
        g=[((PB[t[(k+1)%3],1]-PB[t[(k+2)%3],1])/det,(PB[t[(k+2)%3],0]-PB[t[(k+1)%3],0])/det) for k in range(3)]
        wx=-sum(u[t[k]]*g[k][0] for k in range(3)); wy=-sum(u[t[k]]*g[k][1] for k in range(3))
        body.append((wx,-GAMMA+wy)); area.append(det/2)
    return u,np.array(body),np.array(area)
def boundary_integral(PB,TB,u):  # -∮ u n ds over exposed edges (n outward)
    cnt={}
    for t in TB:
        for k in range(3):
            a,b=t[k],t[(k+1)%3]; key=(min(a,b),max(a,b)); cnt[key]=cnt.get(key,[])+[(a,b)]
    F=np.zeros(2)
    for key,v in cnt.items():
        if len(v)==1:
            a,b=v[0]; dx,dy=PB[b]-PB[a]; n=np.array([dy,-dx])  # outward for CCW (length-weighted)
            F-=0.5*(u[a]+u[b])*n
    return F
Ps,Ts=mesh(0.5); PB,TB=mesh(0.8,0.13)
print(f"seepage mesh {len(Ts)} elems, LA mesh {len(TB)} elems (different meshes)")
for case in ['full','still6','flow']:
    hs,it=seep(Ps,Ts,case)
    u,body,area=apply(PB,TB,Ps,Ts,hs)
    vol=(body*area[:,None]).sum(0); F=boundary_integral(PB,TB,u)
    wsum=((body-np.array([0,-GAMMA]))*area[:,None]).sum(0)
    line=f"[{case}] seep iters={it}  sum A*(-grad u)=({wsum[0]:.6f},{wsum[1]:.4f})  -∮u n ds=({F[0]:.6f},{F[1]:.4f})  |diff|={np.abs(wsum-F).max():.2e}"
    print(line)
    if case=='full':
        print(f"   all elements b'y=-(γ-γw)={-(GAMMA-GW):.4f}: max err {np.abs(body[:,1]+(GAMMA-GW)).max():.2e}, max|b'x| {np.abs(body[:,0]).max():.2e}")
    if case=='still6':
        ys=np.array([PB[t][:,1] for t in TB])
        below=ys.max(1)<=6-1e-9; above=ys.min(1)>=6+1e-9
        print(f"   below WL: max|b'y+(γ-γw)|={np.abs(body[below,1]+(GAMMA-GW)).max():.2e}  above WL: max|b'y+γ|={np.abs(body[above,1]+GAMMA).max():.2e}  straddling={len(TB)-below.sum()-above.sum()}")
        print(f"   buoyancy check: upward force {wsum[1]:.3f} vs γw*A(y<6) = {GW*sum(a for a,t in zip(area,TB) if PB[t][:,1].mean()<6):.3f} (approx, straddling elems)")
