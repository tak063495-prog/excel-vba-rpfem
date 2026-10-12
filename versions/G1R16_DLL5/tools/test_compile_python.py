"""Python dictionary-order reconstruction vs DLL5 complete sparse compiler.
Includes exact canonical signature, old native ordering and reused permutations.
"""
import ctypes as c
import numpy as np
from prepare import ROOT,write,sha

def run():
 dll=c.WinDLL(str(ROOT/'delivery/rpfem_native_x64.dll'));i=c.c_int32;p=c.c_void_p
 schema={'create':[i,i,i,i,i,p,c.POINTER(p)],'destroy':[p],'step':[p,i,c.POINTER(i)],'signature':[p,p,i,c.POINTER(i)],'order_begin':[p,p,p,i],'permute_begin':[p],'counts':[p,p],'export':[p,i,p],'slots':[p,i,i,p]}
 for k,v in schema.items():f=getattr(dll,'rpx_compile_'+k);f.argtypes=v;f.restype=i
 dll.rpx_order_create.argtypes=[i,i,c.POINTER(p)];dll.rpx_order_feed.argtypes=[p,i,i,p,i,p,i];dll.rpx_order_start.argtypes=[p];dll.rpx_order_step.argtypes=[p,i,c.POINTER(i),p,i];dll.rpx_order_destroy.argtypes=[p]
 ptr=lambda a:p(a.ctypes.data);arr=lambda x:np.asarray(x,dtype=np.int32)
 rng=np.random.default_rng(51012);rows=[]
 for seed in range(60):
  nx=int(rng.integers(1,81));nm=seed%8;n=nx+nm;nb=5+seed%27
  cones=[];ip=[0];ids=[]
  for k in range(nb):
   a=list(map(int,rng.choice(nx,size=int(rng.integers(0,min(nx,9)+1)),replace=False)))
   if k%4==3:a=cones[k-2].copy()
   cones.append(a);ids+=a;ip.append(len(ids))
  ep=[0];ei=[];ev=[]
  for k in range(nm):
   a=list(map(int,rng.choice(nx,size=int(rng.integers(1,nx+1)),replace=False)));ei+=a;ev+=list(rng.normal(size=len(a)));ep.append(len(ei))
  cols=[{k:True} for k in range(n)];patterns={};pattern=[]
  for k,a in enumerate(cones):
   key=tuple(a);pattern.append(patterns.get(key,k))
   if key in patterns:continue
   patterns[key]=k
   for j in range(len(a)):
    for l in range(j+1):u,v=sorted([a[l],a[j]]);cols[v][u]=True
  for k in range(nm):
   for a in ei[ep[k]:ep[k+1]]:cols[nx+k][a]=True
  signature=str(n)+':'+';'.join(','.join(map(str,sorted(d))) for d in cols)
  ap0=[0];ai0=[]
  for d in cols:ai0+=list(d);ap0.append(len(ai0))
  ap0=arr(ap0);ai0=arr(ai0);oh=p();perm0=np.zeros(n,np.int32)
  assert dll.rpx_order_create(n,0,c.byref(oh))==0
  try:
   assert dll.rpx_order_feed(oh,0,n,ptr(ap0),len(ap0),ptr(ai0),len(ai0))==0
   assert dll.rpx_order_start(oh)==0;written=i()
   while written.value<n:assert dll.rpx_order_step(oh,64,c.byref(written),ptr(perm0),n)==0
  finally:assert dll.rpx_order_destroy(oh)==0
  for hit in [0,1]:
   perm=perm0.copy() if not hit else rng.permutation(n).astype(np.int32);inv=np.zeros(n,np.int32);inv[perm]=np.arange(n)
   arrays=[arr(ip),arr(ids+[0]),arr(ep),arr(ei+[0]),np.asarray(ev+[0],np.float64)];args=np.array([a.ctypes.data for a in arrays],np.uint64);h=p()
   assert dll.rpx_compile_create(nx,nm,nb,len(ids),len(ei),ptr(args),c.byref(h))==0
   try:
    done=i()
    while done.value!=1:assert dll.rpx_compile_step(h,0,c.byref(done))==0
    length=i();assert dll.rpx_compile_signature(h,None,0,c.byref(length))==0;buffer=c.create_unicode_buffer(length.value+1)
    assert dll.rpx_compile_signature(h,c.cast(buffer,p),length.value,c.byref(length))==0;assert buffer.value==signature
    assert dll.rpx_compile_order_begin(h,ptr(perm),ptr(inv),hit)==0;done.value=0
    while done.value!=3:assert dll.rpx_compile_step(h,0,c.byref(done))==0
    assert dll.rpx_compile_permute_begin(h)==0;done.value=0
    while done.value!=2:assert dll.rpx_compile_step(h,0,c.byref(done))==0
    counts=np.zeros(2,np.int32);assert dll.rpx_compile_counts(h,ptr(counts))==0;nz=int(counts[0])
    outputs=[np.zeros(n,np.int32),np.zeros(n,np.int32),np.zeros(n+1,np.int32),np.zeros(nz,np.int32),np.zeros(n,np.int32),np.zeros(nz),np.zeros(nb,np.int32)];outargs=np.array([a.ctypes.data for a in outputs],np.uint64)
    assert dll.rpx_compile_export(h,nz,ptr(outargs))==0
    assert np.array_equal(outputs[0],perm) and np.array_equal(outputs[1],inv) and np.array_equal(outputs[6],pattern)
    ordered=[{} for _ in range(n)]
    for k,d in enumerate(cols):
     for key in d:u,v=sorted([int(inv[k]),int(inv[key])]);ordered[v][u]=True
    eap=[0];eai=[];diag=np.zeros(n,np.int32)
    for k,d in enumerate(ordered):
     for key in d:d[key]=len(eai);eai.append(key);diag[k]=d[key] if key==k else diag[k]
     eap.append(len(eai))
    base=np.zeros(nz)
    for k in range(nm):
     for at in range(ep[k],ep[k+1]):u,v=sorted([int(inv[nx+k]),int(inv[ei[at]])]);base[ordered[v][u]]+=ev[at]
    assert np.array_equal(outputs[2],eap) and np.array_equal(outputs[3],eai) and np.array_equal(outputs[4],diag) and np.array_equal(outputs[5],base)
    for k,a in enumerate(cones):
     if pattern[k]!=k or not a:continue
     size=len(a);slot=np.zeros((size,size),np.int32,order='F');expected=np.zeros_like(slot)
     for j in range(size):
      for l in range(j+1):u,v=sorted([int(inv[a[j]]),int(inv[a[l]])]);expected[j,l]=ordered[v][u]
     assert dll.rpx_compile_slots(h,k,size*size,ptr(slot))==0;assert np.array_equal(slot,expected)
    rows.append({'seed':seed,'cache_hit':hit,'variables':n,'nnz':nz,'signature_chars':len(signature),'exact':True})
   finally:assert dll.rpx_compile_destroy(h)==0
   assert dll.rpx_compile_destroy(h)==318
 assert dll.rpx_compile_live_handles()==0
 write(ROOT/'results/compile_python.json',{'passed':len(rows),'scope':'Python dictionary-order reconstruction and unchanged native ordering; VBA oracle follows','dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'cases':rows});print('COMPILE PYTHON PASS',len(rows),flush=True)
if __name__=='__main__':run()
