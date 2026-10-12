"""Native Excel failure propagation; hooks exist only in an isolated test copy."""
import gc,json,shutil,ctypes as c
import win32com.client
from support import ROOT,OUTPUT,write
from native_helpers import compile_book
def main():
 folder=ROOT/'results/fan_guards'
 if folder.exists():folder.rename(folder.with_name(folder.name+'_previous_'+str(__import__('time').time_ns())))
 folder.mkdir(exist_ok=False);path=folder/'Guards.xlsm'
 shutil.copyfile(ROOT/'delivery'/OUTPUT,path);shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
 app=win32com.client.DispatchEx('Excel.Application');book=None;rows=[]
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
  cm=book.VBProject.VBComponents('RPX_Mesh').CodeModule;cm.AddFromString((ROOT/'tests/FanGuardWrapper.txt').read_text(encoding='utf8'))
  cm=book.VBProject.VBComponents('RPX_NativeMesh').CodeModule
  text=cm.Lines(1,cm.CountOfLines).replace('Option Explicit','Option Explicit\r\nPublic FanInjected As Long').replace('jobs = jobs + 1','jobs = jobs + 1\r\n    If FanInjected <> 0 Then err.Raise FanInjected, "FanGuardTest", "INJECTED_NATIVE_JOB_FAILURE"')
  cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(text);compile_book(app,book)
  call=lambda n,*a:app.Run("'"+book.Name+"'!"+n,*a)
  original=ROOT/'results/geometry25/P06N_old_vba_mesh.bin'
  for mode in ['vba','auto']:
   call('RPX_NativeSetMode',mode)
   cases=[('cancel',18),('phase',5),('input',5),('capacity',7),('subscript',9)]
   if mode=='auto':cases += [('native_memory',7),('native_unknown',513),('native_cancel',18)]
   for kind,expected in cases:
    msg=str(call('FanGuardTest',kind,str(original)));print(mode,kind,msg,flush=True)
    assert msg.startswith('ERROR '+str(expected)+':') and 'nodes_unchanged=True;elements_unchanged=True'in msg,msg
    rows.append({'mode':mode,'test':kind,'expected_error':expected,'result':msg,'PASS':True})
  write(folder/'summary.json',{'rows':rows,'PASS':True,'scope':'Actual Excel test-only injection, before mesh mutation. Production source contains no hooks.'})
 finally:
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
if __name__=='__main__':main()
