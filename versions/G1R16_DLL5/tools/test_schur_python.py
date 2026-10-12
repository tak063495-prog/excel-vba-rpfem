"""Independent Python matrix fixtures and residuals before the Excel integration."""
import ctypes as C,json
import numpy as np
from prepare import ROOT,write,sha
I=C.c_int32;U=C.c_uint64;P=C.c_size_t;DP=C.POINTER(C.c_double);IP=C.POINTER(I);PP=C.POINTER(P)
def addr(a):return int(a.ctypes.data)
def main():
 dll=C.CDLL(str(ROOT/'delivery/rpfem_native_x64.dll'))
 definitions={'rpx_schur_create':[I]*7+[U,PP,C.POINTER(P)],'rpx_schur_destroy':[P],'rpx_schur_factor_begin':[P,U,U],'rpx_schur_factor_step':[P,U,U,DP,I,IP,IP],'rpx_schur_factor_result':[P,U,U,DP,DP],'rpx_schur_factor_commit':[P,U,U,PP,I],'rpx_schur_solve_begin':[P,U,U,DP,I,I,C.POINTER(U)],'rpx_schur_solve_step':[P,U,U,U,I,IP],'rpx_schur_solve_result':[P,U,U,U,DP,DP],'rpx_schur_live_handles':[],'rpx_numeric':[I]*3+[IP,IP,DP]+[IP]*5+[DP]*4+[IP]}
 for name,args in definitions.items():f=getattr(dll,name);f.argtypes=args;f.restype=I
 rng=np.random.default_rng(160005);maximum=0;count=0;yielded=0
 def ints(v):return np.asarray(v,dtype=np.int32)
 def doubles(v):return np.asarray(v,dtype=np.float64)
 for trial in range(60):
  r=rng.normal(size=(3,3));hess=r@r.T+np.eye(3)*2;coupling=rng.normal(size=(2,3))*.2
  matrix=np.block([[hess,coupling.T],[coupling,np.eye(2)*(-2 if trial%2 else 2)]])
  ids={};values=[]
  for j in range(5):
   for i in range(j+1):ids[i,j]=len(values);values.append(matrix[i,j])
  fullav=doubles(values);fullap=ints([0,1,3,6,10,15]);fullai=ints([i for j in range(5) for i in range(j+1)])
  si=lambda i,j:0 if i==j==0 else 2 if i==j==1 else 1
  layout=[ints([0,1,2]),ints([ids[min(i,j),max(i,j)] for j in range(3) for i in range(3)]),ints([0,2]),ints([0,1]),ints([ids[k,3+i] for i in range(2) for k in range(3)]),ints([0,4]),ints([si(i,j) for j in range(2) for i in range(2)]),ints([0,1]),ints([3,4]),ints([si(i-3,j-3) if i>=3 else -1 for j in range(5) for i in range(j+1)])]
  args=np.asarray([addr(a) for a in layout],dtype=np.uintp);handle=P();assert dll.rpx_schur_create(5,2,1,15,3,2,4,1,args.ctypes.data_as(PP),C.byref(handle))==0
  try:
   assert dll.rpx_schur_factor_begin(handle,1,1)==0
   done=I();reason=I();steps=0
   while not done.value:
    assert dll.rpx_schur_factor_step(handle,1,1,fullav.ctypes.data_as(DP),0,C.byref(done),C.byref(reason))==0;steps+=1;assert steps<100
   assert reason.value==0;schur=doubles([0,0,0]);metrics=np.zeros(12)
   assert dll.rpx_schur_factor_result(handle,1,1,schur.ctypes.data_as(DP),metrics.ctypes.data_as(DP))==0
   expected=matrix[3:,3:]-coupling@np.linalg.solve(hess,coupling.T)
   assert np.allclose(schur,[expected[0,0],expected[0,1],expected[1,1]],rtol=1e-14,atol=1e-14)
   sap=ints([0,1,3]);sai=ints([0,0,1]);cp=ints([0,1,1]);ri=ints([1]);rp=ints([0,0,1]);rn=ints([0]);wa=ints([0]);lv=np.zeros(1);dv=np.zeros(2);scatter=np.zeros(2);minimum=C.c_double(1e300);failure=I(-1)
   assert dll.rpx_numeric(2,0,2,sap.ctypes.data_as(IP),sai.ctypes.data_as(IP),schur.ctypes.data_as(DP),cp.ctypes.data_as(IP),ri.ctypes.data_as(IP),rp.ctypes.data_as(IP),rn.ctypes.data_as(IP),wa.ctypes.data_as(IP),lv.ctypes.data_as(DP),dv.ctypes.data_as(DP),scatter.ctypes.data_as(DP),C.byref(minimum),C.byref(failure))==0
   factor=[fullap,fullai,fullav,cp,ri,lv,dv];fargs=np.asarray([addr(a) for a in factor],dtype=np.uintp)
   assert dll.rpx_schur_factor_commit(handle,1,1,fargs.ctypes.data_as(PP),1)==0
   for check in [0,1]:
    for reuse in [0,1]:
     rhs=rng.normal(size=5);serial=U();done=I();assert dll.rpx_schur_solve_begin(handle,1,1,rhs.ctypes.data_as(DP),check,reuse,C.byref(serial))==0
     assert dll.rpx_schur_solve_step(handle,1,2,serial,0,C.byref(done))==318
     steps=0
     while not done.value:
      assert dll.rpx_schur_solve_step(handle,1,1,serial,0,C.byref(done))==0;steps+=1;assert steps<1000
     yielded+=steps>1;out=np.zeros(5);m=np.zeros(12)
     assert dll.rpx_schur_solve_result(handle,1,1,serial,out.ctypes.data_as(DP),m.ctypes.data_as(DP))==0
     residual=np.linalg.norm(matrix@out-rhs,np.inf)/(np.linalg.norm(matrix,np.inf)*np.linalg.norm(out,np.inf)+np.linalg.norm(rhs,np.inf));maximum=max(maximum,residual);assert residual<1e-14
     assert m[0]==(4 if check else 1) and m[1]==(4 if check else 0) and m[2]==(3 if check else 0)
     count+=1
  finally:assert dll.rpx_schur_destroy(handle)==0
  assert dll.rpx_schur_destroy(handle)==318
 assert dll.rpx_schur_live_handles()==0
 write(ROOT/'results/schur_python.json',{'passed':count,'real_yielded_jobs':yielded,'maximum_independent_backward_error':maximum,'live_handles':0,'dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'validation_type':'Independent Python matrices/residuals; not VBA bitwise comparison'})
 print('Schur Python PASS',count,'max backward error',maximum,flush=True)
if __name__=='__main__':main()
