"""Native geometry job vs Python rule, contracts, cancellation cleanup and lifetime."""
import ctypes as c,json,numpy as np,copy,time
from support import ROOT,write,sha
from diagnose import mesh_from_saved
from fan_repair import repair
from model_support.schema import topology
def main():
 dll=c.WinDLL(str(ROOT/'delivery/rpfem_native_x64.dll'));i=c.c_int32;p=c.c_void_p
 assert dll.rpx_abi()==0x170005
 schema={'create':[i,i,i,p,p,p,p,p,c.POINTER(p)],'step':[p,i,c.POINTER(i)],'result':[p,i,p,p,p,c.POINTER(i)],'destroy':[p],'live_handles':[]}
 for k,args in schema.items():f=getattr(dll,'rpx_meshfan_'+k);f.argtypes=args;f.restype=i
 # Native x87 expressions use the caller's control word; align Python to Excel.
 cw=dll.rpx_test_controlword;cw.argtypes=[i,i];cw.restype=i;oldcw=cw(0,0);cw(1,0x137f)
 m,rec=mesh_from_saved();_,_,edges=topology(m)
 xy=np.ascontiguousarray(m.xy,np.float64);tri=np.ascontiguousarray(m.tri,np.int32)
 eligible=np.asarray(m.materials[m.material_id,0]==0,np.int32);body=np.asarray(np.any(m.body!=0,axis=(1,2)),np.int32)
 ed=np.asarray([[e.elements[0],e.elements[1],*e.loc[0][:2],*e.loc[1][:2],int(e.elements[1]<0 and e.kind=='load' and not np.any(e.loads))]for e in edges],np.int32)
 ptr=lambda a:p(a.ctypes.data);out=[]
 def create(x=xy,t=tri,a=eligible,b=body,g=ed):
  h=p();code=dll.rpx_meshfan_create(len(x),len(t),len(g),ptr(x),ptr(t),ptr(a),ptr(b),ptr(g),c.byref(h));return code,h
 def result(h):
  split=np.full(len(tri),-99,np.int32);cx=np.zeros(len(tri));cy=np.zeros(len(tri));count=i(-1)
  code=dll.rpx_meshfan_result(h,len(tri),ptr(split),ptr(cx),ptr(cy),c.byref(count));return code,split,cx,cy,count.value
 try:
  for label,a,b,g in [('actual',eligible,body,ed),('positive_cohesion',eligible*0,body,ed),('no_body_load',eligible,body*0,ed),('loaded_or_fixed_surface',eligible,body,np.column_stack([ed[:,:6],ed[:,6]*0]).astype(np.int32))]:
   code,h=create(a=a,b=b,g=g);assert code==0
   try:
    assert result(h)[0]==318 # No incomplete plan may escape.
    done=i();steps=0
    while not done.value:
     assert dll.rpx_meshfan_step(h,0,c.byref(done))==0;steps+=1;assert steps<10000
    code,split,cx,cy,count=result(h);assert code==0
    if label=='actual':
     changes=repair(m,rec)[1];assert count==len(changes)==1
     for a in changes:assert split[a['element']]==a['vertex'];assert np.max(np.abs(np.array([cx[a['element']],cy[a['element']]])-a['point']))<1e-14
    else:assert count==0 and np.all(split==-1)
    assert dll.rpx_meshfan_step(h,20,c.byref(done))==318
    out.append({'test':label,'PASS':True,'split_count':count,'steps':steps})
   finally:assert dll.rpx_meshfan_destroy(h)==0
   assert dll.rpx_meshfan_destroy(h)==318
  for label,t,a,g,expected in [('bad_node',tri.copy(),eligible,ed,330),('bad_flag',tri,eligible.copy(),ed,330),('bad_loc',tri,eligible,ed.copy(),330)]:
   if label=='bad_node':t[0,0]=len(xy)
   elif label=='bad_flag':a[0]=2
   else:g[0,2]=3
   code,h=create(t=t,a=a,g=g);assert code==expected and not h.value;out.append({'test':label,'code':code,'PASS':True})
  badxy=xy.copy();badxy[0,0]=np.nan;code,h=create(x=badxy);assert code==206 and not h.value
  out.append({'test':'nonfinite_geometry','code':code,'PASS':True})
  for k in range(100):
   code,h=create();assert code==0;done=i();assert dll.rpx_meshfan_step(h,0,c.byref(done))==0
   assert dll.rpx_meshfan_destroy(h)==0;assert dll.rpx_meshfan_live_handles()==0
  out.append({'test':'100_cancelled_contexts_no_leak','PASS':True})
  assert dll.rpx_meshfan_live_handles()==0
  write(ROOT/'results/native_meshfan.json',{'DLL_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'tests':out,'PASS':True})
  print('NATIVE MESH FAN PASS',len(out),flush=True)
 finally:cw(1,oldcw)
if __name__=='__main__':main()
