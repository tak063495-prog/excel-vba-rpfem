"""Exercise real yields, interleaved contexts, owned RHS and stale guards in DLL4.
Synthetic identity equations; this is a native API test, not an Excel/FEM run.
"""
import ctypes as C,json
import numpy as np
from prepare import ROOT,write,sha
I=C.c_int32;U=C.c_size_t;E=C.c_uint64;P=C.c_void_p

def run():
 dll=C.CDLL(str(ROOT/'delivery/rpfem_native_x64.dll'));records={};n=262144
 idx=np.arange(n,dtype='<i4');ap=np.arange(n+1,dtype='<i4');cp=np.zeros(n+1,dtype='<i4');dummy=np.zeros(1,dtype='<i4');ones=np.ones(n);rhs=idx.astype(float)+.25;out=np.zeros(n)
 address=lambda *arrays:(U*len(arrays))(*(a.ctypes.data for a in arrays))
 def f(name,args):
  fn=getattr(dll,'rpx_'+name);fn.argtypes=args;fn.restype=I;return fn
 create=f('refine_create',[I,C.POINTER(U)]);begin=f('refine_begin',[U,I,E,P]);step=f('refine_step',[U,I,E,P,I,C.POINTER(I)]);result=f('refine_result',[U,I,E,P,C.POINTER(I),C.POINTER(I),C.POINTER(C.c_double)]);destroy=f('refine_destroy',[U]);live=f('refine_live_handles',[])
 args=address(ap,idx,ones,cp,dummy,ones,ones);h=U();other=U();done=I();checks=I();corrections=I();res=C.c_double()
 assert create(n,C.byref(h))==create(n,C.byref(other))==0 and h.value!=other.value
 owned=rhs.copy();assert begin(h,n,101,owned.ctypes.data)==0;owned.fill(-999)
 assert begin(h,n,102,rhs.ctypes.data)==318
 assert step(h,n,100,args,1,C.byref(done))==318 and result(h,n,101,out.ctypes.data,C.byref(checks),C.byref(corrections),C.byref(res))==318
 assert step(h,n,101,args,1,C.byref(done))==0 and done.value==0,'test must actually yield'
 assert begin(other,n,201,rhs.ctypes.data)==0
 assert step(other,n,201,args,1,C.byref(done))==0 and done.value==0
 old=other.value;assert destroy(other)==0 and live()==1 and create(n,C.byref(other))==0 and other.value!=old and destroy(old)==318
 assert destroy(other)==0
 steps=1
 while not done.value:
  assert step(h,n,101,args,1,C.byref(done))==0;steps+=1
 assert result(h,n,101,out.ctypes.data,C.byref(checks),C.byref(corrections),C.byref(res))==0
 assert np.array_equal(out,rhs) and checks.value==1 and corrections.value==0 and res.value==0
 assert result(h,n,100,out.ctypes.data,C.byref(checks),C.byref(corrections),C.byref(res))==318 and destroy(h)==0 and live()==0
 records['ordinary']={'n':n,'steps':steps,'actual_yield':True,'owned_rhs':True,'interleaved_contexts':True,'destroy_partly_processed_context':True,'wrong_epoch_reentry_early_result_and_stale_id_rejected':True,'exact_identity_solution':True,'live_handles_after':0}
 create=f('hlinear_create',[I,I,I,I,C.POINTER(U)]);begin=f('hlinear_begin',[U,E,I,I,I,P]);step=f('hlinear_step',[U,E,I,P,P,U,I,C.POINTER(I)]);result=f('hlinear_result',[U,E,P,P,P,P]);destroy=f('hlinear_destroy',[U]);live=f('hlinear_live_handles',[])
 # One scalar cone and one unit G coefficient per variable; original H operator = I.
 dims=np.ones(n,dtype='<i4');co=np.zeros((n,4));co[:,0]=co[:,2]=1.;hh=np.zeros((9,n,4));hh[0,:,0]=hh[0,:,2]=1.;scale_dd=np.zeros((n,4));scale_dd[:,0]=scale_dd[:,2]=1.;reg=np.zeros(4)
 op=address(idx,scale_dd,dims,idx,ap,idx,idx,co,hh,dummy,dummy,reg,reg,reg);factor=address(cp,dummy,ones,ones,ones)
 initial=address(rhs,idx,scale_dd);best=np.zeros((n,2));bdd=np.zeros((n,2));scaled=np.zeros(n);metrics=np.zeros(14)
 assert create(n,n,n,n,C.byref(h))==create(n,n,n,n,C.byref(other))==0
 assert begin(h,301,42,0,0,initial)==0 and begin(h,302,42,0,0,initial)==318
 assert step(h,300,42,op,factor,0,1,C.byref(done))==318 and step(h,301,43,op,factor,0,1,C.byref(done))==318
 assert result(h,301,best.ctypes.data,scaled.ctypes.data,bdd.ctypes.data,metrics.ctypes.data)==318
 assert step(h,301,42,op,factor,0,1,C.byref(done))==0 and done.value==0
 assert begin(other,401,43,0,0,initial)==0 and step(other,401,43,op,factor,0,1,C.byref(done))==0 and done.value==0
 old=other.value;assert destroy(other)==0 and live()==1 and create(n,n,n,n,C.byref(other))==0 and other.value!=old and destroy(old)==318
 assert destroy(other)==0
 steps=1
 while not done.value:
  assert step(h,301,42,op,factor,0,1,C.byref(done))==0;steps+=1
 assert result(h,301,best.ctypes.data,scaled.ctypes.data,bdd.ctypes.data,metrics.ctypes.data)==0
 assert np.array_equal(best[:,0],rhs) and np.all(best[:,1]==0) and np.array_equal(scaled,rhs) and metrics[0]==1 and metrics[1]==0 and metrics[2]==0
 assert destroy(h)==0 and live()==0
 records['HSD']={'n':n,'steps':steps,'actual_yield':True,'interleaved_contexts':True,'destroy_partly_processed_context':True,'wrong_epoch_factor_reentry_early_result_and_stale_id_rejected':True,'exact_DD_identity_solution':True,'live_handles_after':0}
 records['dll_sha256']=sha(ROOT/'delivery/rpfem_native_x64.dll');records['pass']=True;write(ROOT/'results/job_lifecycle.json',records);print(json.dumps(records),flush=True)
if __name__=='__main__':run()
