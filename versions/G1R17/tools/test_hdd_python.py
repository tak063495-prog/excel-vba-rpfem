"""DD KKT assembly independently checked at 80-digit precision.
Mixed scalar/Lorentz cones, repeated KKT slots, regularization and scaling.
"""
import ctypes as c
from decimal import Decimal,localcontext
import numpy as np
from support import ROOT,write,sha

def run():
 dll=c.WinDLL(str(ROOT/'delivery/rpfem_native_x64.dll'));i=c.c_int32;p=c.c_void_p
 dll.rpx_hdd_step.argtypes=[i,i,i,i,i,c.POINTER(i),p,i];dll.rpx_hdd_step.restype=i
 def coef(a):
  a=np.asarray(a,np.float64).reshape(-1);out=np.zeros((len(a),4));out[:,0]=a;bits=a.view(np.uint64).copy()&np.uint64(0xfffffffff8000000);out[:,2]=bits.view(np.float64);out[:,3]=a-out[:,2];return out
 rng=np.random.default_rng(51005);maximum=Decimal(0);rows=[];D=Decimal.from_float
 with localcontext() as ctx:
  ctx.prec=80
  for seed in range(60):
   n=5;nx=3;nb=3;dims=np.array([3,1,3],np.int32);ip=np.array([0,3,6,9],np.int32);sp=np.array([0,6,12,18],np.int32);ap=np.array([0,1,3,6,10,15],np.int32);ai=np.array([r for col in range(n) for r in range(col+1)],np.int32);diag=np.array([int(ap[j+1])-1 for j in range(n)],np.int32)
   slot=np.array([int(ap[j])+k for _ in range(nb) for j in range(nx) for k in range(j+1)],np.int32);at=np.array([0,9,12],np.int32)
   H=np.empty((nb,3,3));G=[]
   for k in range(nb):
    T=rng.normal(size=(3,3));H[k]=T@T.T+np.eye(3);G.append(rng.normal(size=(int(dims[k]),nx)))
   coeff=coef(np.concatenate([g.reshape(-1) for g in G]));hh=coef(H.transpose(2,1,0).reshape(-1));scales=rng.uniform(.2,2,n);scale=coef(scales)
   base=rng.normal(size=15);out=np.zeros((15,2));perm=np.arange(n,dtype=np.int32)
   arrays=[base,out,dims,ip,sp,slot,at,coeff,hh,diag,perm,scale,ap,ai];args=np.array([a.ctypes.data for a in arrays],np.uint64)
   for phase,total in [(0,15),(1,3),(2,5),(3,5)]:
    cursor=i()
    while cursor.value<total:assert dll.rpx_hdd_step(n,nx,nb,15,phase,c.byref(cursor),p(args.ctypes.data),0)==0
   expected=list(map(D,base))
   for k in range(nb):
    g=G[k];dim=int(dims[k])
    for j in range(nx):
     for l in range(j+1):expected[int(ap[j])+l]+=sum(D(float(g[r,l]))*D(float(H[k,r,q]))*D(float(g[q,j])) for r in range(dim) for q in range(dim))
   for j in range(n):expected[int(diag[j])]=expected[int(diag[j])]+D(1e-10) if j<nx else D(-1e-10)
   for col in range(n):
    for at0 in range(int(ap[col]),int(ap[col+1])):expected[at0]*=D(float(scales[col]))*D(float(scales[int(ai[at0])]))
   error=max(abs(D(float(out[k,0]))+D(float(out[k,1]))-expected[k])/(1+abs(expected[k])) for k in range(15));assert error<Decimal('2e-29'),(seed,error);maximum=max(maximum,error);rows.append({'seed':seed,'relative_error':str(error)})
 result={'passed':len(rows),'max_decimal_relative_error':str(maximum),'decimal_precision':80,'scope':'independent DD assembly verification; not VBA bitwise proof','dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'cases':rows}
 write(ROOT/'results/hdd_python.json',result);print('HDD PYTHON PASS',len(rows),maximum,flush=True)
if __name__=='__main__':run()
