"""Bitwise old DLL4 vs fused DLL5 DD arithmetic, including error precedence.
Separate saved DLL binaries are the oracle; no replicated optimized formula.
"""
import ctypes as c,struct,math,time,random
from prepare import ROOT,DLL4,write,sha
class DD(c.Structure):_fields_=[('hi',c.c_double),('lo',c.c_double)]
def run():
 paths=[DLL4/'delivery/rpfem_native_x64.dll',ROOT/'delivery/rpfem_native_x64.dll'];dlls=[c.WinDLL(str(p)) for p in paths]
 assert [d.rpx_abi() for d in dlls]==[0x160004,0x160005]
 funcs=[]
 for d in dlls:
  f=d.rpx_dd_test;f.argtypes=[c.c_int32,c.POINTER(DD),c.POINTER(DD),c.POINTER(DD)];f.restype=c.c_int32;funcs.append(f)
 rng=random.Random(51005);cases=[]
 for _ in range(1800):
  a=rng.uniform(-2,2)*10.**rng.randint(-150,150);b=rng.uniform(-2,2)*10.**rng.randint(-150,150)
  cases.append((DD(a,a*rng.uniform(-1,1)*1e-16),DD(b,b*rng.uniform(-1,1)*1e-16)))
 for a in [0.,-0.,1.,-1.,1e-300,-1e-300,1e300,-1e300,math.inf,-math.inf,math.nan]:
  for b in [0.,-0.,1.,-1.,1e-300,1e300,math.inf,math.nan]:cases.append((DD(a,0),DD(b,0)))
 counts={};t=time.perf_counter()
 for op in range(4):
  for index,(a,b) in enumerate(cases):
   out=[DD(),DD()];codes=[f(op,c.byref(a),c.byref(b),c.byref(o)) for f,o in zip(funcs,out)]
   assert codes[0]==codes[1],(op,index,codes)
   if codes[0]==0:assert bytes(out[0])==bytes(out[1]),(op,index,bytes(out[0]).hex(),bytes(out[1]).hex(),(a.hi,a.lo,b.hi,b.lo))
   counts[str(codes[0])]=counts.get(str(codes[0]),0)+1
 result={'cases':len(cases)*4,'bitwise_exact_successes_and_identical_error_precedence':True,'codes':counts,'seconds':time.perf_counter()-t,'DLL4_sha256':sha(paths[0]),'DLL5_sha256':sha(paths[1]),'scope':'native primitive binary oracle; actual VBA/FEM verification follows'}
 write(ROOT/'results/dd_core_python.json',result);print('DD CORE PASS',result['cases'],counts,flush=True)
if __name__=='__main__':run()
