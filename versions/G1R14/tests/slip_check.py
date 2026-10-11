# Python mirror of DrawSlip selection + TraceSlipSurface + Kasa fit (synthetic circular slip, self-test case 4)
import math, heapq
nx,ny=40,20; xy=[(0.5*i,0.5*j) for j in range(ny+1) for i in range(nx+1)]
tri=[]
for j in range(ny):
    for i in range(nx):
        a=j*(nx+1)+i;b=a+1;c=b+nx+1;d=a+nx+1; tri+=[(a,b,c),(a,c,d)]
inside=[(sum(xy[v][0] for v in t)/3-16)**2+(sum(xy[v][1] for v in t)/3-16)**2<121 for t in tri]
edges={}
for e,t in enumerate(tri):
    for k in range(3):
        a,b=t[k],t[(k+1)%3]; edges.setdefault((min(a,b),max(a,b)),[]).append(e)
E=list(edges.items()); rn=len(xy)
def vel(p,e): return (-(p[1]-16),p[0]-16) if inside[e] else (0,0)
jump=[];length=[]
for (a,b),els in E:
    L=math.dist(xy[a],xy[b]); length.append(L)
    if len(els)<2: jump.append(0);continue
    m=0
    for p in (xy[a],xy[b],((xy[a][0]+xy[b][0])/2,(xy[a][1]+xy[b][1])/2)):
        v1=vel(p,els[0]);v2=vel(p,els[1]); m=max(m,math.hypot(v1[0]-v2[0],v1[1]-v2[1]))
    jump.append(m)
n=sum(1 for (k,els) in E if len(els)==2); t=max(jump)
nzv=sorted(j for j in jump if j>1e-6*t); ref=nzv[int(0.95*(len(nzv)-1))]
cands=[i for i in range(len(E)) if jump[i]>=0.5*ref]; maxCount=int(0.1*n)+1
score=[jump[i]*length[i] for i in range(len(E))]
cut=sorted(score[i] for i in cands)[len(cands)-maxCount] if len(cands)>maxCount else 0
sel=[jump[i]>=0.5*ref and score[i]>=cut for i in range(len(E))]
adj=[[] for _ in range(rn)]
for i,((a,b),els) in enumerate(E):
    w=length[i]/(jump[i]/ref+0.001); adj[a].append((b,w)); adj[b].append((a,w))
surf=[False]*rn; selN=[False]*rn; sl=[]
for i,((a,b),els) in enumerate(E):
    if len(els)==1: surf[a]=surf[b]=True
    if sel[i]: selN[a]=selN[b]=True; sl.append(length[i])
R=2.5*sum(sl)/len(sl)
cand=[v for v in range(rn) if surf[v] and (selN[v] or any(selN[w] for w,_ in adj[v]) or any(selN[w] and math.dist(xy[v],xy[w])<=R for w in range(rn)))]
if len(cand)>30: cand=[cand[int(i*(len(cand)-1)/29)] for i in range(30)]
def sp(s):
    d=[1e300]*rn;pr=[-1]*rn;d[s]=0;h=[(0,s)]
    while h:
        dv,v=heapq.heappop(h)
        if dv>d[v]: continue
        for w,c in adj[v]:
            if dv+c<d[w]: d[w]=dv+c;pr[w]=v;heapq.heappush(h,(d[w],w))
    return d,pr
pairs=[]
for i in range(len(cand)-1):
    d,pr=sp(cand[i])
    for j in range(i+1,len(cand)):
        w=cand[j];pl=0;inter=0;v=w
        while pr[v]>=0: pl+=math.dist(xy[v],xy[pr[v]]); v=pr[v]; inter+= (not surf[v])
        if pl>0 and inter>0: pairs.append((cand[i],w,d[w]/pl,math.dist(xy[cand[i]],xy[w])))
ms=max(p[3] for p in pairs); best=min(p[2] for p in pairs if p[3]>=0.25*ms)
pick=max((p for p in pairs if p[3]>=0.25*ms and p[2]<=1.5*best),key=lambda p:p[3])
d,pr=sp(pick[0]); path=[];v=pick[1]
while v>=0: path.append(v); v=pr[v]
X=[xy[v] for v in path]
import numpy as np
A=np.array([[x,y,1] for x,y in X]); z=np.array([-(x*x+y*y) for x,y in X]); D,Ee,F=np.linalg.lstsq(A,z,rcond=None)[0]
cx,cy=-D/2,-Ee/2; r=math.sqrt(cx*cx+cy*cy-F)
print(f"selected={sum(sel)} cand={len(cand)} pairs={len(pairs)} chain={len(path)} ends={xy[path[0]]},{xy[path[-1]]} surf={surf[path[0]]},{surf[path[-1]]} circle=({cx:.2f},{cy:.2f}) r={r:.2f}")
