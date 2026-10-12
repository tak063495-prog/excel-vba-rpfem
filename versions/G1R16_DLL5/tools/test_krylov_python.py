"""Independent dense-system checks of the DLL5 GMRES before VBA integration.
The operator is G' H G with one equality; numeric DD residual remains authoritative.
"""
import ctypes as c,json,time
import numpy as np
from prepare import ROOT,write,sha

def run():
 dll=c.WinDLL(str(ROOT/'delivery/rpfem_native_x64.dll'))
 i=c.c_int32;u=c.c_uint64;p=c.c_void_p
 api={
 'create':[i,i,i,i,c.POINTER(p)],'destroy':[p],
 'begin':[p,u,i,i,i,p,p,c.c_double],
 'step':[p,u,i,p,p,p,i,c.POINTER(i)],'result':[p,u,p,p],
 'bytes':[p,c.POINTER(u)]}
 for name,args in api.items():
  f=getattr(dll,'rpx_krylov_'+name);f.argtypes=args;f.restype=i
 ptr=lambda a:p(a.ctypes.data)
 def coef(values):
  v=np.asarray(values,dtype=np.float64).reshape(-1);out=np.zeros((len(v),4));out[:,0]=v
  bits=v.view(np.uint64).copy();bits&=np.uint64(0xfffffffff8000000);out[:,2]=bits.view(np.float64);out[:,3]=v-out[:,2];return out
 rng=np.random.default_rng(51212);rows=[];maximum=0
 for seed in range(40):
  nx=3;n=4;ns=3;nb=1
  G=rng.normal(size=(3,3))*.2+np.eye(3);T=rng.normal(size=(3,3));H=T@T.T+np.eye(3)
  E=rng.normal(size=3);scale=10**rng.uniform(-1,1,n)
  physical=np.zeros((n,n));physical[:3,:3]=G.T@H@G+np.eye(3)*1e-10;physical[3,:3]=E;physical[:3,3]=E;physical[3,3]=-1e-10
  A=scale[:,None]*physical*scale[None,:]
  arrays=[np.arange(n,dtype=np.int32),coef(scale),np.array([3],np.int32),np.array([0],np.int32),np.array([0,3],np.int32),np.arange(3,dtype=np.int32),np.array([0],np.int32),coef(G.reshape(-1)),coef(H.flatten(order='F')),np.array([0,3],np.int32),np.arange(3,dtype=np.int32),coef(E),coef([1e-10]),coef([-1e-10])]
  op=np.array([a.ctypes.data for a in arrays],dtype=np.uint64)
  cp=np.zeros(n+1,np.int32);ri=np.zeros(1,np.int32);lv=np.zeros(1);diag=np.diag(A).copy();diag[-1]=-max(1.,np.max(np.abs(diag)))
  # Exact-zero initial answer, reproducible ordinary rhs; solve same system using numpy.
  b=rng.normal(size=n);bdd=np.column_stack((b,np.zeros(n)));answer=np.zeros((n,2))
  expected=np.linalg.solve(A,b)
  for kind in [0,1]:
   dv=np.column_stack((diag,np.zeros(n))) if kind else diag;lf=np.zeros((1,2)) if kind else lv
   fa=np.array([cp.ctypes.data,ri.ctypes.data,lf.ctypes.data,dv.ctypes.data,scale.ctypes.data],np.uint64)
   handle=p();assert dll.rpx_krylov_create(n,nx,nb,ns,c.byref(handle))==0
   try:
    assert dll.rpx_krylov_begin(handle,1,7,kind,0,ptr(bdd),ptr(answer),1e300)==0
    done=i();calls=0;t=time.perf_counter()
    while not done.value:
     assert dll.rpx_krylov_step(handle,1,7,ptr(op),ptr(fa),None,0,c.byref(done))==0
     calls+=1;assert calls<20000
    out=np.zeros((n,2));m=np.zeros(4);assert dll.rpx_krylov_result(handle,1,ptr(out),ptr(m))==0
    x=out[:,0]+out[:,1];res=float(np.max(np.abs(A@x-b)));err=res/(float(np.linalg.norm(A,np.inf)*np.linalg.norm(x,np.inf)+np.linalg.norm(b,np.inf)))
    forward=float(np.max(np.abs(x-expected))/(1+np.max(np.abs(expected))))
    assert m[2]==1 and m[0]<=1e-11,(seed,kind,m,res)
    assert err<2e-14 and forward<2e-11,(seed,kind,err,forward)
    assert dll.rpx_krylov_result(handle,2,ptr(out),ptr(m))==318
    maximum=max(maximum,err);rows.append({'seed':seed,'kind':kind,'calls':calls,'steps':int(m[1]),'backward_error':err,'forward_error':forward,'seconds':time.perf_counter()-t})
   finally:assert dll.rpx_krylov_destroy(handle)==0
   assert dll.rpx_krylov_destroy(handle)==318
 assert dll.rpx_krylov_live_handles()==0
 result={'passed':len(rows),'max_independent_backward_error':maximum,'scope':'independent numpy dense solve; not VBA bitwise proof','dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'jobs':rows}
 write(ROOT/'results/krylov_python.json',result);print('KRYLOV PYTHON PASS',len(rows),maximum,flush=True)
if __name__=='__main__':run()
