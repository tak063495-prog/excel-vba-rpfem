import gc,shutil,time
import win32com.client
from prepare import ROOT,write,sha
from native_helpers import compile_book
def main():
 folder=ROOT/'results/guard_oracle';folder.mkdir(exist_ok=True)
 path=folder/'Guard_Oracle.xlsm';shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm',path);shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
 app=win32com.client.DispatchEx('Excel.Application');book=None
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
  book.VBProject.VBComponents('RPX_Guard').CodeModule.AddFromString((ROOT/'tests/GuardOracle.txt').read_text(encoding='utf8'))
  book.VBProject.VBComponents('RPX_Solve').CodeModule.AddFromString((ROOT/'tests/LiveIdentityOracle.txt').read_text(encoding='utf8'))
  compile_book(app,book);call=lambda name,*args:app.Run("'"+book.Name+"'!"+name,*args)
  call('RPX_NativeSetMode','auto');assert str(call('RPX_NativeInfo')).startswith('DLL_READY')
  t=time.perf_counter();msg=str(call('DLL5GuardOracle'));assert msg.startswith('PASS'),msg
  live=str(call('DLL5LiveIdentityOracle'));assert live.startswith('PASS'),live
  write(folder/'summary.json',{'message':msg,'live_identity':live,'seconds':time.perf_counter()-t,'source_sha256':sha(path),'dll_sha256':sha(folder/'rpfem_native_x64.dll'),'actual_Excel':True,'security_unmodified':app.AutomationSecurity})
  print(msg,flush=True)
  print(live,flush=True)
 finally:
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
if __name__=='__main__':main()
