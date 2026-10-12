"""Actual Excel fixed-Fs solves on the repaired coarse mesh, then export for Python."""
import gc,json,shutil,time,struct
import win32com.client
from support import ROOT,BASE,OUTPUT,REPO,write,sha
from native_helpers import compile_book

def inject(book,name):
 text=(ROOT/'tests'/(name+'.bas')).read_text(encoding='utf8');text='\r\n'.join(s for s in text.splitlines() if not s.startswith('Attribute VB_'))
 c=book.VBProject.VBComponents.Add(1);c.Name=name;c.CodeModule.AddFromString(text)

def main():
 folder=ROOT/'results/fixed_native';folder.mkdir(exist_ok=False)
 p=folder/'Fixed.xlsm';shutil.copyfile(ROOT/'delivery'/OUTPUT,p);shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
 app=win32com.client.DispatchEx('Excel.Application');book=None;rows=[]
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(p),UpdateLinks=0)
  for name in ['HSDRegression','MeshProbe']:inject(book,name)
  compile_book(app,book);call=lambda n,*a:app.Run("'"+book.Name+"'!"+n,*a);call('RPX_NativeSetMode','auto')
  for side,fs in [('lower',.5),('upper',1.)]:
   prefix=folder/(side+'_'+str(fs));print('START',side,fs,flush=True);t=time.perf_counter();msg=str(call('FixedFem',side,fs,str(prefix)))
   r={'side':side,'Fs':fs,'seconds':time.perf_counter()-t,'message':msg};rows.append(r);write(folder/'summary.json',{'rows':rows,'DLL_SHA256':sha(folder/'rpfem_native_x64.dll')});print(msg,flush=True);assert msg.startswith('PASS'),msg
   data=prefix.with_name(prefix.name+'_mesh.bin').read_bytes();nn,ne=struct.unpack_from('<ii',data);assert nn==147 and ne==245
  book.Save();write(folder/'summary.json',{'rows':rows,'DLL_SHA256':sha(folder/'rpfem_native_x64.dll'),'passed':True})
 finally:
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
if __name__=='__main__':main()
