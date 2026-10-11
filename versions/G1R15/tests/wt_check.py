# Reference checks for water-table modes (mirrors RPX_SeepWT.bas)
import sys, math, numpy as np
sys.path.insert(0,'/workspace/limitvba'); import seep_check as sc

def wt_eval(tx,ty,x):
    """y_wt(x) and sin^2(theta): piecewise linear, constant beyond ends, theta per segment."""
    n=len(tx)
    if x<=tx[0]: return ty[0], (0.0 if n<2 else sin2(tx[0],ty[0],tx[1],ty[1]))
    if x>=tx[-1]: return ty[-1], (0.0 if n<2 else sin2(tx[-2],ty[-2],tx[-1],ty[-1]))
    for i in range(n-1):
        if x<=tx[i+1]:
            t=(x-tx[i])/(tx[i+1]-tx[i]); return ty[i]+t*(ty[i+1]-ty[i]), sin2(tx[i],ty[i],tx[i+1],ty[i+1])
def sin2(x0,y0,x1,y1):
    dx=x1-x0; dy=y1-y0; return dy*dy/(dx*dx+dy*dy)
def wt_head(tx,ty,x,y,mode):
    yw,s2=wt_eval(tx,ty,x)
    return yw if mode==2 else yw-(yw-y)*s2

def check_mode1():
    th=math.radians(30); gam=19.; gw=9.81; c=5.; phi=math.radians(30); H=6.; dwt=2.  # table 2 m (vertical) below surface
    # slab mesh: surface y = -tan(th) x, x in [0,40], vertical depth 0..H (sheared grid)
    nx,ny=80,24; xy=[];idx={}
    for j in range(ny+1):
        for i in range(nx+1):
            x=40*i/nx; ys=-math.tan(th)*x; idx[i,j]=len(xy); xy.append((x,ys-H*(1-j/ny)))
    tri=[]
    for j in range(ny):
        for i in range(nx):
            a,b,cc,d=idx[i,j],idx[i+1,j],idx[i+1,j+1],idx[i,j+1]; tri+=[(a,b,cc),(a,cc,d)]
    tx=[-10,50]; ty=[-math.tan(th)*-10-dwt, -math.tan(th)*50-dwt]
    h=np.array([wt_head(tx,ty,x,y,1) for x,y in xy]); Y=np.array([p[1] for p in xy])
    u=np.maximum(gw*(h-Y),0)
    hw=np.array([max(-math.tan(th)*x-dwt-y,0) for x,y in xy])
    err_u=np.max(np.abs(u-gw*hw*math.cos(th)**2))
    # integrate effective body force along slope-normal through middle column -> stresses on plane at vertical depth H
    bx=np.zeros(len(tri)); by=np.zeros(len(tri)); A=np.zeros(len(tri)); C=np.zeros((len(tri),2))
    for e,(a,b,cc) in enumerate(tri):
        ar2,bb,cv=sc.tri_geom(xy,a,b,cc); A[e]=ar2/2
        gx=sum(u[v]*bb[k] for k,v in enumerate((a,b,cc)))/ar2; gy=sum(u[v]*cv[k] for k,v in enumerate((a,b,cc)))/ar2
        bx[e]=-gx; by[e]=-gam-gy; C[e]=np.mean([xy[a],xy[b],xy[cc]],axis=0)
    # infinite-slope: per unit plan length, stresses on plane = (sum of b' over column)/(width). use central column x in [15,25]
    sel=(C[:,0]>15)&(C[:,0]<25); W=10.0
    Fx=(bx[sel]*A[sel]).sum()/W; Fy=(by[sel]*A[sel]).sum()/W   # force per horizontal length
    n=np.array([math.sin(th),math.cos(th)]); t=np.array([math.cos(th),-math.sin(th)])  # outward normal of base (pointing up into slab), downslope tangent
    # base plane length per horizontal length = 1/cos
    sig=-(Fx*n[0]+Fy*n[1])*math.cos(th); tau=(Fx*t[0]+Fy*t[1])*math.cos(th)
    fs_num=(c+sig*math.tan(phi))/tau
    hwb=H-dwt; ub=gw*hwb*math.cos(th)**2
    fs_an=(c+(gam*H*math.cos(th)**2-ub)*math.tan(phi))/(gam*H*math.sin(th)*math.cos(th))
    print(f"mode1 infinite slope: max|u-γw hw cos²θ|={err_u:.2e} kPa ; FS mesh={fs_num:.5f} analytic={fs_an:.5f} diff={100*(fs_num-fs_an)/fs_an:.3f}%")
    return abs(fs_num-fs_an)/fs_an<0.005 and err_u<1e-9

# ---- inverse
def build_dam(H1,nx=40,ny=20,L=20.,Ht=10.,H2=2.):
    xy=[];idx={}
    for j in range(ny+1):
        for i in range(nx+1): idx[i,j]=len(xy); xy.append((L*i/nx,Ht*j/ny))
    tri=[]
    for j in range(ny):
        for i in range(nx):
            a,b,c,d=idx[i,j],idx[i+1,j],idx[i+1,j+1],idx[i,j+1]; tri+=[(a,b,c),(a,c,d)]
    n=len(xy); bt=[0]*n; bv=[0.0]*n
    for j in range(ny+1):          # upstream "水位" kind: y<=H1 fixed head, above impermeable
        a=idx[0,j]
        if xy[a][1]<=H1+1e-9: bt[a]=1; bv[a]=H1
        b=idx[nx,j]
        if xy[b][1]<=H2+1e-9: bt[b]=1; bv[b]=H2
        else: bt[b]=2
    return xy,tri,bt,bv
def fs_segments(xy,tri,h):
    seg=[]
    for t in tri:
        pts=[]
        for i in range(3):
            a=t[i]; b=t[(i+1)%3]; pa=h[a]-xy[a][1]; pb=h[b]-xy[b][1]
            if (pa>=0 and pb<0) or (pa<0 and pb>=0):
                s=pa/(pa-pb); pts.append((xy[a][0]+s*(xy[b][0]-xy[a][0]), xy[a][1]+s*(xy[b][1]-xy[a][1])))
        if len(pts)==2: seg.append(pts[0]+pts[1])
    return seg
def fs_y(seg,x):
    best=None; near=None; nd=1e300
    for x1,y1,x2,y2 in seg:
        lo,hi=min(x1,x2),max(x1,x2)
        if hi-lo>1e-12 and lo-1e-9<=x<=hi+1e-9:
            y=y1+(y2-y1)*(x-x1)/(x2-x1)
            if best is None or y>best: best=y
        for xe,ye in ((x1,y1),(x2,y2)):
            if abs(xe-x)<nd: nd=abs(xe-x); near=ye
    return best if best is not None else near
def forward(H1,obsx):
    xy,tri,bt,bv=build_dam(H1); h,q,it,kr=sc.solve(xy,tri,[0]*len(tri),[1.0],[1.0],bt,bv)
    seg=fs_segments(xy,tri,h); return np.array([fs_y(seg,x) for x in obsx])
def lm(p0,obsx,obsy,lo,hi,maxeval=40):
    p=np.array(p0,float); nev=0
    r=forward(p[0],obsx)-obsy; nev+=1; lam=1e-3
    while nev<maxeval:
        J=np.zeros((len(obsx),len(p)))
        for k in range(len(p)):
            d=max(1e-3*abs(p[k]),0.01); q=p.copy(); q[k]+=d; J[:,k]=(forward(q[0],obsx)-obsy-r)/d; nev+=1
        g=J.T@r; A=J.T@J
        step=np.linalg.solve(A+lam*np.diag(np.diag(A)+1e-12),-g)
        q=np.clip(p+step,lo,hi); rq=forward(q[0],obsx)-obsy; nev+=1
        if rq@rq<r@r: p,r=q,rq; lam=max(lam/10,1e-7)
        else: lam*=10
        if np.max(np.abs(step))<1e-4: break
    return p,math.sqrt(r@r/len(r)),nev,np.linalg.cond(A)
def check_mode3():
    obsx=np.array([2.,5.,8.,11.,14.])
    obsy=forward(10.0,obsx)
    p,rms,nev,cond=lm([8.0],obsx,obsy,[0.],[20.])
    print(f"mode3 dam: observed y={np.round(obsy,3)}; start 8.0 -> H1={p[0]:.4f} (err {100*(p[0]-10)/10:.3f}%), RMS={rms:.2e} m, evals={nev}, cond={cond:.2g}")
    return abs(p[0]-10)<0.1
if __name__=="__main__":
    a=check_mode1(); b=check_mode3(); print("PASS" if a and b else "FAIL")
