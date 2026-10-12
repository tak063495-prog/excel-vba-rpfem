"""Issue #6 regression (Python mirror of the corrected RPX_SeepSolve convergence test).

Convergence = unrelaxed fixed-point change < tol AND scaled free-node flux residual < rtol
AND no seepage-face switching. Mirrors versions/G1R15*/src/vba/RPX_Seep.bas (G1R15 #6 fix).
Run: python seep_convergence_check.py   (needs versions/G1R14/tests/seep_check.py)
"""
import sys, pathlib, importlib.util
here = pathlib.Path(__file__).resolve()
sys.path.insert(0, str(here.parents[2] / 'G1R14' / 'tests'))
import seep_check as base  # original solver (dmax-only criterion)
src = (here.parents[2] / 'G1R14' / 'tests' / 'seep_check.py').read_text(encoding='utf-8')
import math
from seep_check import rcm, tri_geom
CONV=[False]
LAST=None
def solve_new(xy, tri, mat, kx, ky, bctype, bcval, krmin=1e-4, beta=20.0, relax=0.5, tol=1e-6, maxit=300, verbose=False, rtol=1e-6):
    CONV[0]=False
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
        dmax=0.0; dfix=0.0
        for i in range(n):
            hn=x[perm[i]]
            if not fixed[i]:
                dfix=max(dfix,abs(hn-h[i])); hn=h[i]+relax*(hn-h[i])
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
        # residual on free nodes with the accepted heads and kr re-evaluated at accepted heads
        kr2=[]; q2=[0.0]*n
        for e,t in enumerate(tri):
            area2,bb,cc=tri_geom(xy,*t); p=sum(h[v]-xy[v][1] for v in t)/3
            k=1.0 if p>=0 else max(krmin, math.exp(beta*p)); kr2.append(k); s_=k/(2*area2)
            for i in range(3):
                for j in range(3):
                    q2[t[i]]+=s_*(kx[mat[e]]*bb[i]*bb[j]+ky[mat[e]]*cc[i]*cc[j])*h[t[j]]
        fixedn=[bctype[i]==1 or active[i] for i in range(n)]
        qs=sum(abs(q2[i]) for i in range(n) if fixedn[i]) + sum(abs(bcval[i]) for i in range(n) if bctype[i]==3)
        qs=max(qs,1e-30)
        res=max([abs(q2[i]-(bcval[i] if bctype[i]==3 else 0.0)) for i in range(n) if not fixedn[i]] or [0.0])/qs
        global LAST; LAST=(it,dfix,res,changed)
        if dfix<tol and res<rtol and changed==0: CONV[0]=True; break
    return h,q,it,kr


def four_node(relax):
    xy=[(0,0),(1,0),(1,1),(0,1)]; tri=[(0,1,2),(0,2,3)]
    return xy,tri,[0,0],[1.0],[1.0],[1,1,0,0],[10.0,9.0,0.0,0.0]

if __name__=='__main__':
    ok=True
    for r in (1e-8,0.5,1.0):
        h,q,it,kr=base.solve(*four_node(r),relax=r)
        h2,q2,it2,kr2=solve_new(*four_node(r),relax=r)
        good=abs(h2[2]-28/3)<1e-5 and abs(h2[3]-29/3)<1e-5
        print(f'relax={r:g} old: it={it} h2={h[2]:.9f} | new: it={it2} converged={CONV[0]} h2={h2[2]:.9f} last={LAST}')
        if CONV[0] and not good: ok=False   # never accept a wrong field
        if r>=0.5 and not CONV[0]: ok=False
    base_solve=base.solve; base.solve=solve_new; Q=base.dam(); base.solve=base_solve
    print('dam (new):',Q[:3],'converged',CONV[0])
    ok = ok and CONV[0] and abs(Q[0]-2.4180720)<1e-5
    print('PASS' if ok else 'FAIL'); sys.exit(0 if ok else 1)
