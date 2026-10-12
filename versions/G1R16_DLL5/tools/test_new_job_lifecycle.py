"""Actual one-ms yields and stale/owned-state checks for DLL5 new contexts.
Synthetic identity matrices; deliberately not represented as FEM/Fs acceptance.
"""
import ctypes as c
import numpy as np
from prepare import ROOT,write,sha
I=c.c_int32;U=c.c_uint64;P=c.c_void_p
def run():
 dll=c.WinDLL(str(ROOT/'delivery/rpfem_native_x64.dll'));records={};ptr=lambda a:P(a.ctypes.data);addresses=lambda *a:np.array([x.ctypes.data for x in a],np.uint64)
 def f(name,args):fn=getattr(dll,'rpx_'+name);fn.argtypes=args;fn.restype=I;return fn
 create=f('schur_create',[I,I,I,I,I,I,I,U,P,c.POINTER(P)]);destroy=f('schur_destroy',[P]);begin=f('schur_factor_begin',[P,U,U]);step=f('schur_factor_step',[P,U,U,P,I,c.POINTER(I),c.POINTER(I)]);result=f('schur_factor_result',[P,U,U,P,P]);commit=f('schur_factor_commit',[P,U,U,P,I]);sb=f('schur_solve_begin',[P,U,U,P,I,I,c.POINTER(U)]);ss=f('schur_solve_step',[P,U,U,U,I,c.POINTER(I)]);sr=f('schur_solve_result',[P,U,U,U,P,P])
 nb=32768;nf=3*nb+1;pos=np.arange(3*nb,dtype=np.int32);hp=np.full((nb,9),-1,np.int32)
 for k in range(3):hp[:,k+3*k]=np.arange(nb)*3+k
 ip=np.arange(nb+1,dtype=np.int32);rows=np.zeros(nb,np.int32);bp=np.full(3*nb,-1,np.int32);slots=np.zeros(nb,np.int32);inv=np.zeros(1,np.int32);full=np.array([nf-1],np.int32);direct=np.full(nf,-1,np.int32);direct[-1]=0
 a=addresses(pos,hp,ip,rows,bp,ip,slots,inv,full,direct);matrix=np.ones(nf);h=P();other=P();done=I();reason=I()
 assert create(nf,1,nb,nf,1,nb,nb,1,ptr(a),c.byref(h))==0
 assert begin(h,1,1)==0 and begin(h,1,2)==318
 assert step(h,1,2,ptr(matrix),1,c.byref(done),c.byref(reason))==318
 assert step(h,1,1,ptr(matrix),1,c.byref(done),c.byref(reason))==0 and done.value==0
 steps=1
 while not done.value:assert step(h,1,1,ptr(matrix),1,c.byref(done),c.byref(reason))==0;steps+=1
 av=np.zeros(1);metrics=np.zeros(12);assert result(h,1,1,ptr(av),ptr(metrics))==0 and av[0]==1
 ap=np.arange(nf+1,dtype=np.int32);ai=np.arange(nf,dtype=np.int32);cp=np.zeros(2,np.int32);dummy=np.zeros(1,np.int32);empty=np.zeros(1);diag=np.ones(1);fa=addresses(ap,ai,matrix,cp,dummy,empty,diag)
 assert commit(h,1,1,ptr(fa),0)==0
 rhs=np.arange(nf,dtype=float)+.25;original=rhs.copy();serial=U();assert sb(h,1,1,ptr(rhs),1,1,c.byref(serial))==0;rhs.fill(-999)
 assert ss(h,1,1,serial.value,1,c.byref(done))==0 and done.value==0
 solve_steps=1
 while not done.value:assert ss(h,1,1,serial.value,1,c.byref(done))==0;solve_steps+=1
 out=np.zeros(nf);assert sr(h,1,1,serial.value,ptr(out),ptr(metrics))==0 and np.array_equal(out,original) and metrics[3]==0 and metrics[2]==3
 old=h.value;assert destroy(h)==0 and destroy(old)==318
 assert create(nf,1,nb,nf,1,nb,nb,1,ptr(a),c.byref(other))==0 and other.value!=old
 assert begin(other,1,1)==0 and step(other,1,1,ptr(matrix),1,c.byref(done),c.byref(reason))==0 and done.value==0 and destroy(other)==0
 assert dll.rpx_schur_live_handles()==0
 records['Schur']={'variables':nf,'factor_steps':steps,'checked_solve_steps':solve_steps,'actual_yield':True,'owned_rhs':True,'stale_generation_and_handle_reentry_rejected':True,'partly_processed_context_destroyed':True,'three_corrections_four_checks':True,'exact_identity_solution':True,'live_handles':0}
 create=f('krylov_create',[I,I,I,I,c.POINTER(P)]);destroy=f('krylov_destroy',[P]);begin=f('krylov_begin',[P,U,I,I,I,P,P,c.c_double]);step=f('krylov_step',[P,U,I,P,P,P,I,c.POINTER(I)]);result=f('krylov_result',[P,U,P,P])
 n=65536;idx=np.arange(n,dtype=np.int32);cp=np.zeros(n+1,np.int32);ip=np.arange(n+1,dtype=np.int32);dims=np.ones(n,np.int32);ones=np.ones(n);coeff=np.zeros((n,4));coeff[:,0]=coeff[:,2]=1;hh=np.zeros((9,n,4));hh[0,:,0]=hh[0,:,2]=1;reg=np.zeros(4)
 op=addresses(idx,coeff,dims,idx,ip,idx,idx,coeff,hh,dummy,dummy,reg,reg,reg);fa=addresses(cp,dummy,empty,ones,ones);bdd=np.column_stack((idx.astype(float)+.25,np.zeros(n)));answer=np.zeros((n,2));h=P();other=P()
 assert create(n,n,n,n,c.byref(h))==create(n,n,n,n,c.byref(other))==0
 assert begin(h,101,7,0,0,ptr(bdd),ptr(answer),1e300)==0 and begin(h,102,7,0,0,ptr(bdd),ptr(answer),1e300)==318
 expected=bdd[:,0].copy();bdd.fill(-999)
 assert step(h,100,7,ptr(op),ptr(fa),None,1,c.byref(done))==318 and step(h,101,8,ptr(op),ptr(fa),None,1,c.byref(done))==318
 assert step(h,101,7,ptr(op),ptr(fa),None,1,c.byref(done))==0 and done.value==0
 assert begin(other,201,8,0,0,ptr(bdd),ptr(answer),1e300)==0 and step(other,201,8,ptr(op),ptr(fa),None,1,c.byref(done))==0 and done.value==0 and destroy(other)==0
 steps=1
 while not done.value:assert step(h,101,7,ptr(op),ptr(fa),None,1,c.byref(done))==0;steps+=1
 m=np.zeros(4);assert result(h,101,ptr(answer),ptr(m))==0 and m[2]==1 and m[0]<=1e-11
 assert np.max(np.abs(answer[:,0]+answer[:,1]-expected))<1e-10
 old=h.value;assert destroy(h)==0 and destroy(old)==318 and dll.rpx_krylov_live_handles()==0
 records['Krylov']={'variables':n,'steps':steps,'actual_yield':True,'owned_rhs_and_answer':True,'interleaved_contexts':True,'partly_processed_context_destroyed':True,'wrong_epoch_factor_reentry_and_stale_id_rejected':True,'accepted_original_DD_gate':True,'residual':float(m[0]),'live_handles':0}
 write(ROOT/'results/new_job_lifecycle.json',{'pass':True,'records':records,'dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll')});print('NEW JOB LIFECYCLE PASS',records,flush=True)
if __name__=='__main__':run()
