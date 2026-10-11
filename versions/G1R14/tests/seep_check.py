# Line-for-line port of RPX_Seep.bas (T3 steady seepage, kr-reduction free surface)
import math, sys
import numpy as np

def tri_geom(xy, a, b, c):
    x1,y1=xy[a];x2,y2=xy[b];x3,y3=xy[c]
    area2=(x2-x1)*(y3-y1)-(x3-x1)*(y2-y1)
    bb=[y2-y3,y3-y1,y1-y2]; cc=[x3-x2,x1-x3,x2-x1]
    return area2, bb, cc

def rcm(n, tri):
    adj=[set() for _ in range(n)]
    for t in tri:
        for i in range(3):
            for j in range(3):
                if i!=j: adj[t[i]].add(t[j])
    deg=[len(a) for a in adj]; order=[]; seen=[False]*n
    while len(order)<n:
        s=min((i for i in range(n) if not seen[i]), key=lambda i:deg[i])
        seen[s]=True; q=[s]; h=0
        while h<len(q):
            v=q[h]; h+=1
            for w in sorted(adj[v], key=lambda i:deg[i]):
                if not seen[w]: seen[w]=True; q.append(w)
        order+=q
    order.reverse(); perm=[0]*n
    for k,v in enumerate(order): perm[v]=k
    return perm  # perm[node]=equation

def solve(xy, tri, mat, kx, ky, bctype, bcval, krmin=1e-4, beta=20.0, relax=0.5, tol=1e-6, maxit=300, verbose=False):
    n=len(xy); m=len(tri)
    perm=rcm(n,tri)
    fc=list(range(n))
    for t in tri:
        p=[perm[v] for v in t]; lo=min(p)
        for v in p: fc[v]=min(fc[v],lo)
    ptr=[0]*(n+1)
    for i in range(n): ptr[i+1]=ptr[i]+(i-fc[i]+1)
    h=[0.0]*n
    hfix=max(bcval[i] for i in range(n) if bctype[i]==1)
    for i in range(n): h[i]=bcval[i] if bctype[i]==1 else hfix
    active=[bctype[i]==2 for i in range(n)]  # seepage face starts fixed h=y
    for i in range(n):
        if active[i]: h[i]=xy[i][1]
    for it in range(1,maxit+1):
        fixed=[bctype[i]==1 or active[i] for i in range(n)]
        fval=[bcval[i] if bctype[i]==1 else xy[i][1] for i in range(n)]
        a=[0.0]*ptr[n]; f=[0.0]*n
        for i in range(n):
            if bctype[i]==3: f[perm[i]]+=bcval[i]
        kr=[0.0]*m
        for e,t in enumerate(tri):
            area2,bb,cc=tri_geom(xy,*t)
            p=sum(h[v]-xy[v][1] for v in t)/3
            kr[e]=1.0 if p>=0 else max(krmin, math.exp(beta*p))
            s=kr[e]/(2*area2)
            for i in range(3):
                gi=t[i]; ri=perm[gi]
                if fixed[gi]: continue
                for j in range(3):
                    gj=t[j]; rj=perm[gj]
                    kij=s*(kx[mat[e]]*bb[i]*bb[j]+ky[mat[e]]*cc[i]*cc[j])
                    if fixed[gj]: f[ri]-=kij*fval[gj]
                    elif rj<=ri: a[ptr[ri]+rj-fc[ri]]+=kij
        for i in range(n):
            if fixed[i]:
                ri=perm[i]; a[ptr[ri]+ri-fc[ri]]=1.0; f[ri]=fval[i]
        # skyline Cholesky
        for i in range(n):
            for j in range(fc[i],i+1):
                s=a[ptr[i]+j-fc[i]]
                for k in range(max(fc[i],fc[j]),j):
                    s-=a[ptr[i]+k-fc[i]]*a[ptr[j]+k-fc[j]]
                if j<i: a[ptr[i]+j-fc[i]]=s/a[ptr[j]+j-fc[j]]
                else:
                    if s<=0: raise RuntimeError("SEEP_NOT_SPD")
                    a[ptr[i]+i-fc[i]]=math.sqrt(s)
        x=f[:]
        for i in range(n):
            s=x[i]
            for k in range(fc[i],i): s-=a[ptr[i]+k-fc[i]]*x[k]
            x[i]=s/a[ptr[i]+i-fc[i]]
        for i in range(n-1,-1,-1):
            x[i]/=a[ptr[i]+i-fc[i]]
            for k in range(fc[i],i): x[k]-=a[ptr[i]+k-fc[i]]*x[i]
        dmax=0.0
        for i in range(n):
            hn=x[perm[i]]
            if not fixed[i]: hn=h[i]+relax*(hn-h[i])
            dmax=max(dmax,abs(hn-h[i])); h[i]=hn
        # reactions (outflow positive = -K h at node); seepage-face switching
        q=[0.0]*n
        for e,t in enumerate(tri):
            area2,bb,cc=tri_geom(xy,*t); s=kr[e]/(2*area2)
            for i in range(3):
                for j in range(3):
                    q[t[i]]+=s*(kx[mat[e]]*bb[i]*bb[j]+ky[mat[e]]*cc[i]*cc[j])*h[t[j]]
        changed=0
        qtol=1e-9*max(abs(v) for v in q)
        for i in range(n):
            if bctype[i]==2:
                if active[i] and q[i]>qtol: active[i]=False; changed+=1   # q>0 = inflow -> release
                elif (not active[i]) and h[i]>xy[i][1]: active[i]=True; h[i]=xy[i][1]; changed+=1
        if verbose: print(it,dmax,changed)
        if dmax<tol and changed==0: break
    return h,q,it,kr

def dam(L=20.,H1=10.,H2=2.,nx=40,ny=20,**kw):
    xy=[];idx={}
    for j in range(ny+1):
        for i in range(nx+1):
            idx[i,j]=len(xy); xy.append((L*i/nx,H1*j/ny))
    tri=[]
    for j in range(ny):
        for i in range(nx):
            a,b,c,d=idx[i,j],idx[i+1,j],idx[i+1,j+1],idx[i,j+1]
            tri+=[(a,b,c),(a,c,d)]
    n=len(xy); bt=[0]*n; bv=[0.0]*n
    for j in range(ny+1):
        bt[idx[0,j]]=1; bv[idx[0,j]]=H1
        y=xy[idx[nx,j]][1]
        if y<=H2+1e-9: bt[idx[nx,j]]=1; bv[idx[nx,j]]=H2
        else: bt[idx[nx,j]]=2
    h,q,it,kr=solve(xy,tri,[0]*len(tri),[1.0],[1.0],bt,bv,**kw)
    Q=sum(q[idx[0,j]] for j in range(ny+1))   # inflow upstream (positive = K h)
    Qout=-sum(q[idx[nx,j]] for j in range(ny+1))
    # free surface: x -> y where p=0 along each column
    fs=[]
    for i in range(nx+1):
        ys=None
        for j in range(ny):
            p0=h[idx[i,j]]-xy[idx[i,j]][1]; p1=h[idx[i,j+1]]-xy[idx[i,j+1]][1]
            if p0>=0 and p1<0: ys=xy[idx[i,j]][1]+(xy[idx[i,j+1]][1]-xy[idx[i,j]][1])*p0/(p0-p1); break
        if ys is None: ys=H1
        fs.append((L*i/nx,ys))
    face=max([xy[idx[nx,j]][1] for j in range(ny+1) if bt[idx[nx,j]]==2 and q[idx[nx,j]]< -1e-6]+[H2])
    return Q,Qout,it,fs,face

if __name__=="__main__":
    Qd=(10**2-2**2)/(2*20)
    for nx,ny in [(20,10),(40,20)]:
        Q,Qo,it,fs,face=dam(nx=nx,ny=ny)
        print(f"mesh {nx}x{ny}: iters={it} Qin={Q:.4f} Qout={Qo:.4f} Dupuit={Qd:.4f} err={100*(Q-Qd)/Qd:.2f}% seepface_top={face:.3f}")
    for x,y in fs[::5]:
        print(f"x={x:5.1f} FEM={y:6.3f} Dupuit={math.sqrt(100-(96)*x/20):6.3f}")
