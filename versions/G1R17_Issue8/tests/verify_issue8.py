"""Python 3 standard library; binary64 mirror of G1R17 RPX_Seep.bas.
RCM, skyline, element/assembly order, Picard and active-set switching retained.
Only convergence normalization changes in mode=mixed. Not Excel execution.
"""
import math, json, sys
EPS=sys.float_info.epsilon

def ordering(n,tri):
    adj=[set() for _ in range(n)]
    for t in tri:
        for i in t: adj[i].update(j for j in t if i!=j)
    seen=set(); order=[]
    while len(order)<n:
        root=min((i for i in range(n) if i not in seen),key=lambda i:(len(adj[i]),i))
        seen.add(root); q=[root]
        for v in q:
            for j in sorted(adj[v],key=lambda i:(len(adj[i]),i)):
                if j not in seen: seen.add(j); q.append(j)
        order+=q
    perm=[0]*n
    for i,v in enumerate(reversed(order)): perm[v]=i
    fc=list(range(n))
    for t in tri:
        a=min(perm[v] for v in t)
        for v in t: fc[perm[v]]=min(fc[perm[v]],a)
    ptr=[0]
    for i in range(n): ptr.append(ptr[-1]+i-fc[i]+1)
    return perm,fc,ptr

def skyline(a,x,fc,ptr):
    n=len(x)
    for i in range(n):
        for j in range(fc[i],i+1):
            s=a[ptr[i]+j-fc[i]]
            for k in range(max(fc[i],fc[j]),j): s-=a[ptr[i]+k-fc[i]]*a[ptr[j]+k-fc[j]]
            a[ptr[i]+j-fc[i]]=s/a[ptr[j]+j-fc[j]] if j<i else math.sqrt(s)
    for i in range(n):
        s=x[i]
        for k in range(fc[i],i): s-=a[ptr[i]+k-fc[i]]*x[k]
        x[i]=s/a[ptr[i]+i-fc[i]]
    for i in reversed(range(n)):
        x[i]/=a[ptr[i]+i-fc[i]]
        for k in range(fc[i],i): x[k]-=a[ptr[i]+k-fc[i]]*x[i]
    return x

def matrices(xy,tri,h,kx,ky,beta=20.,krmin=1e-4):
    out=[]
    for t in tri:
        (x1,y1),(x2,y2),(x3,y3)=[xy[v] for v in t]
        area=(x2-x1)*(y3-y1)-(x3-x1)*(y2-y1)
        assert area>0
        b=[y2-y3,y3-y1,y1-y2]; c=[x3-x2,x1-x3,x2-x1]
        p=0.
        for v in t: p+=(h[v]-xy[v][1])/3.
        kr=1. if p>=0 else max(krmin,math.exp(beta*p))
        s=kr/(2.*area)
        out.append([[s*(kx*b[i]*b[j]+ky*c[i]*c[j]) for j in range(3)] for i in range(3)])
    return out

def flux(n,tri,ks,h):
    q=[0.]*n; mag=[0.]*n; count=[0]*n
    for t,K in zip(tri,ks):
        for i,gi in enumerate(t):
            for j,gj in enumerate(t):
                term=K[i][j]*h[gj]; q[gi]+=term; mag[gi]+=abs(term); count[gi]+=1
    return q,mag,count

def metrics(q,mag,count,bc,val,active,tol=1e-6):
    qs=0.
    for i in range(len(q)):
        if bc[i]==1 or active[i]: qs+=abs(q[i])
        if bc[i]==3: qs+=abs(val[i])
    old=mixed=backward=rmax=0.
    for i in range(len(q)):
        if bc[i]==1 or active[i]: continue
        f=val[i] if bc[i]==3 else 0.; r=abs(q[i]-f); rmax=max(rmax,r)
        # Conservative engineering roundoff allowance, NOT rigorous geometric error bound.
        neps=(count[i]+32)*EPS; gamma=neps/(1-neps)
        roundoff=gamma*(mag[i]+abs(f))
        limit=tol*qs+roundoff
        old=max(old,r/max(qs,1e-300))
        mixed=max(mixed,r/limit if limit>0 else (0. if r==0 else math.inf))
        backward=max(backward,r/max(mag[i]+abs(f),1e-300))
    return old,mixed,backward,rmax,qs

def solve(xy,tri,bc,val,kx=1.,ky=1.,relax=.5,mode='mixed',tol=1e-6,htol=1e-6,maxit=300,initial=None):
    n=len(xy); perm,fc,ptr=ordering(n,tri)
    hf=max(val[i] for i in range(n) if bc[i]==1)
    active=[v==2 for v in bc]
    h=[val[i] if bc[i]==1 else (xy[i][1] if active[i] else hf) for i in range(n)] if initial is None else initial[:]
    switches=0
    for it in range(1,maxit+1):
        fixed=[bc[i]==1 or active[i] for i in range(n)]
        fv=[val[i] if bc[i]==1 else xy[i][1] for i in range(n)]
        a=[0.]*ptr[n]; f=[0.]*n; ks=matrices(xy,tri,h,kx,ky)
        for i in range(n):
            if bc[i]==3: f[perm[i]]+=val[i]
        for t,K in zip(tri,ks):
            for i,gi in enumerate(t):
                ri=perm[gi]
                if not fixed[gi]:
                    for j,gj in enumerate(t):
                        rj=perm[gj]; kij=K[i][j]
                        if fixed[gj]: f[ri]-=kij*fv[gj]
                        elif rj<=ri: a[ptr[ri]+rj-fc[ri]]+=kij
        for i in range(n):
            if fixed[i]: ri=perm[i]; a[ptr[ri]+ri-fc[ri]]=1.; f[ri]=fv[i]
        x=skyline(a,f,fc,ptr); dfix=delta=0.
        for i in range(n):
            hn=x[perm[i]]
            if not fixed[i]: dfix=max(dfix,abs(hn-h[i])); hn=h[i]+relax*(hn-h[i])
            delta=max(delta,abs(hn-h[i])); h[i]=hn
        q,_,_=flux(n,tri,ks,h); qtol=1e-9*max(map(abs,q)); changed=0
        for i in range(n):
            if bc[i]==2:
                if active[i] and q[i]>qtol: active[i]=False; changed+=1
                elif not active[i] and h[i]>xy[i][1]: active[i]=True; h[i]=xy[i][1]; changed+=1
        switches+=changed
        qr,mag,count=flux(n,tri,matrices(xy,tri,h,kx,ky),h)
        old,mixed,backward,rmax,qs=metrics(qr,mag,count,bc,val,active,tol)
        passed=(old<tol if mode=='old' else mixed<1)
        ok=dfix<htol and passed and changed==0
        if ok: break
    return dict(ok=ok,it=it,h=h,dfix=dfix,delta=delta,old=old,mixed=mixed,backward=backward,rmax=rmax,qs=qs,switches=switches,active=active,q=qr,sflux=q)

def dam(nx=8,ny=4):
    xy=[(20.*i/nx,10.*j/ny) for j in range(ny+1) for i in range(nx+1)]; tri=[]
    for j in range(ny):
        for i in range(nx):
            a=j*(nx+1)+i; b=a+1; c=b+nx+1; d=a+nx+1
            tri.extend([(a,b,c),(a,c,d)])
    bc=[0]*len(xy); val=[0.]*len(xy)
    for j in range(ny+1):
        a=j*(nx+1); bc[a]=1; val[a]=10.; b=a+nx
        bc[b]=1 if xy[b][1]<=2.+1e-9 else 2; val[b]=2. if bc[b]==1 else 0.
    return xy,tri,bc,val

def main():
    tests={}; hydro=([(0.,0.),(7.,0.),(1.,3.)],[(0,1,2)],[1,1,0],[10.,10.,0.])
    square=([(0.,0.),(1.,0.),(1.,1.),(0.,1.)],[(0,1,2),(0,2,3)],[1,1,0,0],[10.,9.,0.,0.])
    for mode in ('old','mixed'): tests['hydro_'+mode]=solve(*hydro,mode=mode)
    assert not tests['hydro_old']['ok'] and tests['hydro_old']['old']==8.
    assert tests['hydro_mixed']['ok'] and tests['hydro_mixed']['it']==1
    for relax in (.5,1e-8):
        r=solve(*square,relax=relax); tests['issue6_'+str(relax)]=r
        assert r['ok']==(relax==.5)
    r=solve(*square,relax=1e-8,initial=[10.,9.,28/3,29/3]); tests['exact_small_relax']=r; assert r['ok']
    # independent scales, offset includes elevations so pressure/kr unchanged
    scale_count=0
    for k in (1e-12,1.,1e12):
      for length in (1e-3,1.,1e3):
       for offset in (0.,1e3,1e6):
        xy,t,bc,val=square
        xy=[(x*length,y*length+offset) for x,y in xy]
        val=[v*length+offset if b==1 else 0. for b,v in zip(bc,val)]
        r=solve(xy,t,bc,val,kx=k,ky=.13*k,htol=1e-6*length)
        assert r['ok'],(k,length,offset,r)
        scale_count+=1
    tests['scaled_pass_count']=scale_count
    for k in (1e-12,1.,1e12):
        tests['tiny_relax_k'+str(k)]=solve(*square,kx=k,ky=k,relax=1e-8)
        assert not tests['tiny_relax_k'+str(k)]['ok']
    # Nonzero nodal flux and very small imposed gradients.
    tests['nodal_flux']=solve(hydro[0],hydro[1],[1,1,3],[10.,10.,1e-8])
    assert tests['nodal_flux']['ok']
    for drop in (1e-3,1e-6,1e-9):
        r=solve(square[0],square[1],square[2],[10.,10.-drop,0.,0.]); tests['drop_'+str(drop)]=r; assert r['ok']
    for shape in ((8,4),(12,6),(20,10)):
        key='dam_'+str(shape)
        for mode in ('old','mixed'): tests[key+'_'+mode]=solve(*dam(*shape),mode=mode)
        a,b=tests[key+'_old'],tests[key+'_mixed']
        assert a['ok']==b['ok']==(shape!=(8,4))
        assert a['h']==b['h'] and a['it']==b['it']
    hydro_count=0
    for shape in (hydro,square):
      xy,t,bc,val=shape
      for k in (1e-12,1.,1e12):
       for datum in (-10.,10.,1e6):
        for turn in range(3):
         tt=[tuple(list(e)[turn:]+list(e)[:turn]) for e in t]
         vv=[datum if b==1 else 0. for b in bc]
         r=solve(xy,tt,bc,vv,kx=k,ky=.13*k)
         assert r['ok'],('hydro',k,datum,turn,r)
         hydro_count+=1
    tests['hydro_variant_pass_count']=hydro_count
    # Perturb accepted solution: mixed rejects even if update tolerance alone passes.
    drop=1e-3; h=[10.,10.-drop,10.-2*drop/3+5e-7,10.-drop/3]
    q,mag,c=flux(4,square[1],matrices(square[0],square[1],h,1.,1.),h)
    tests['perturbed_metrics']=metrics(q,mag,c,square[2],[10.,10.-drop,0.,0.],[False]*4)
    assert tests['perturbed_metrics'][1]>1
    r=solve(square[0],square[1],square[2],[10.,10.-drop,0.,0.],initial=h,maxit=1)
    tests['perturbed_one_iteration']=r
    assert r['dfix']<1e-6 and r['mixed']>1 and not r['ok']
    # Compare naive backward error at a large datum.
    offset=1e6; h=[v+offset for v in [10.,9.,28/3+.1,29/3]]
    xy=[(x,y+offset) for x,y in square[0]]; val=[1000010.,1000009.,0.,0.]
    q,mag,c=flux(4,square[1],matrices(xy,square[1],h,1.,1.),h)
    tests['large_datum_imbalance']=metrics(q,mag,c,square[2],val,[False]*4)
    assert tests['large_datum_imbalance'][1]>1 and tests['large_datum_imbalance'][2]<1e-6
    xy=[(0.,0.),(6.961343695493474,0.),(.43948462081667794,2.1986849333421516)]
    for mode in ('old','mixed'):
        r=solve(xy,[(0,1,2)],[1,0,2],[xy[2][1],0.,0.],mode=mode,maxit=20)
        tests['face_roundoff_'+mode]=r
        assert not r['ok'] and r['switches']==20
    # No arbitrary denominator floor: zero budget passes only exact zero residual.
    assert metrics([0.], [0.], [0], [0], [0.], [False])[1] == 0.
    assert math.isinf(metrics([1.], [0.], [0], [0], [0.], [False])[1])
    tests['zero_budget_checks'] = 2
    print(json.dumps(tests,indent=2))
if __name__=='__main__': main()
