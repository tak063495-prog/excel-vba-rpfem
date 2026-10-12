"""Native Excel reopen/compile/buttons/DLL check on both finished presets."""
import gc,json
import win32com.client
from support import ROOT,write,sha
from native_helpers import compile_book
def main():
 records=[]
 for path in sorted((ROOT/'delivery').glob('*.xlsm')):
  before=sha(path);app=win32com.client.DispatchEx('Excel.Application');book=None
  try:
   app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0,ReadOnly=True)
   compile_book(app,book);assert book.VBProject.VBComponents.Count==71
   info=str(app.Run("'"+book.Name+"'!RPX_NativeInfo"));assert info.startswith('DLL_READY'),info
   actions=[]
   for sh in book.Worksheets:
    for shape in sh.Shapes:
     action=shape.OnAction
     if action:actions.append({'sheet':sh.Name,'name':shape.Name,'action':action})
   assert any(a['action'].endswith('RPX_Run')for a in actions)
   assert not any('!'in a['action']and book.Name not in a['action']for a in actions),actions
   records.append({'file':path.name,'modules':71,'compile':'PASS','DLL':info,'buttons':actions,'Fs_lower':book.Worksheets('解析結果').Range('B4').Value2,'Fs_upper':book.Worksheets('解析結果').Range('C4').Value2,'AutomationSecurity_unmodified':app.AutomationSecurity,'read_only':True})
  finally:
   if book is not None:book.Close(False)
   app.Quit();book=None;app=None;gc.collect()
  assert sha(path)==before
 write(ROOT/'results/reopen_delivery.json',{'books':records,'PASS':True});print('REOPEN COMPILE BUTTONS DLL PASS',len(records),flush=True)
if __name__=='__main__':main()
