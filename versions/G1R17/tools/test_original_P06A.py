"""One old DLL5 fixed-Fs P06A coarse solve, no Fs search or adaptation."""
import gc,json,shutil,time
import win32com.client
from support import ROOT,BASE,SOURCE,write,sha,configure
from fixed_native import inject
from native_helpers import compile_book
def main():
 folder=ROOT/'results/original_P06A_fixed';folder.mkdir(exist_ok=False);path=folder/'Original.xlsm'
 shutil.copyfile(SOURCE,path);shutil.copyfile(BASE/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
 app=win32com.client.DispatchEx('Excel.Application');book=None
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
  configure(book,json.loads((ROOT/'data/inputs/P06A.json').read_text(encoding='utf8')))
  for n in ['HSDRegression','MeshProbe']:inject(book,n)
  compile_book(app,book);call=lambda n,*a:app.Run("'"+book.Name+"'!"+n,*a);call('RPX_NativeSetMode','auto')
  t=time.perf_counter();msg=str(call('FixedFem','lower',.5,str(folder/'P06A_original')));assert msg.startswith('PASS'),msg
  write(folder/'summary.json',{'case':'P06A','Fs':.5,'side':'lower','result':msg,'seconds':time.perf_counter()-t,'DLL_sha256':sha(folder/'rpfem_native_x64.dll'),'scope':'Original243 mesh; one fixed point only; full/Fs search not run'})
  print(msg,flush=True)
 finally:
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
if __name__=='__main__':main()
