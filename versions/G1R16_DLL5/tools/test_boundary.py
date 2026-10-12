import gc,json,shutil,struct
import numpy as np
import win32com.client
from prepare import ROOT,write,sha
from test_hsd_regression import inject
from native_helpers import compile_book
def run():
 folder=ROOT/'results/boundary_oracle';folder.mkdir(parents=True,exist_ok=True)
 rng=np.random.default_rng(16004);cases=[]
 for exponent in [-100,-30,-10,0,10,30,100]:
  for _ in range(240):
   scale=10.**exponent;s=rng.normal(size=2)*scale;rad=np.linalg.norm(s)
   for margin in [1e-3,1e-9,1e-14]:cases.append((np.r_[rad*(1+margin),s],rng.normal(size=3)*scale))
 for d0 in [0.,1.,-1.,1e-30,-1e-30]:
  for s0 in [-1.,0.,1e-30,1.,1e30]:cases.append((np.asarray([s0]),np.asarray([d0])))
 for x,d in [([2,1,0],[1,1,0]),([2,1,0],[-1,-1,0]),([1,1,0],[0,0,0]),([2,0,0],[-2,0,0]),([2,1,0],[0,1,0])]:cases.append((np.asarray(x,dtype=float),np.asarray(d,dtype=float)))
 parts=[struct.pack('<i',len(cases))]
 for x,d in cases:parts.extend([struct.pack('<i',len(x)),x.astype('<f8').tobytes(),d.astype('<f8').tobytes()])
 (folder/'inputs.bin').write_bytes(b''.join(parts))
 path=folder/'Boundary.xlsm';shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm',path);shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
 app=win32com.client.DispatchEx('Excel.Application');book=None
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0);inject(book,'BoundaryOracle');compile_book(app,book)
  for mode in ['vba','auto']:print(app.Run("'"+book.Name+"'!DumpBoundary",str(folder/'inputs.bin'),str(folder/(mode+'.bin')),mode),flush=True)
 finally:
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
 a=(folder/'vba.bin').read_bytes();b=(folder/'auto.bin').read_bytes();assert len(a)==len(b)==4+18*len(cases)
 differences=[i for i in range(len(cases)) if a[4+18*i:4+18*(i+1)]!=b[4+18*i:4+18*(i+1)]]
 write(folder/'summary.json',{'cases':len(cases),'bitwise_exact':not differences,'differences':differences[:30],'dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll')})
 print('Boundary',len(cases),'cases; differences:',differences[:30],flush=True);assert not differences
if __name__=='__main__':run()
