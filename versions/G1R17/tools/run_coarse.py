"""Only TARGET200 q8 ADAPT0 P06N; never starts the full adaptive case."""
import gc,json,shutil,time,threading,traceback,argparse
import win32com.client,win32process,win32gui,win32con
from support import ROOT,BASE,REPO,OUTPUT,write,sha,configure

def main():
 parser=argparse.ArgumentParser();parser.add_argument('--case',choices=['P06N','P06A'],default='P06N');args=parser.parse_args()
 case=args.case;folder=ROOT/'results'/('coarse_g1r17_'+case);folder.mkdir(exist_ok=False);p=folder/OUTPUT.replace('P06N',case)
 shutil.copyfile(ROOT/'delivery'/OUTPUT,p);shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
 app=win32com.client.DispatchEx('Excel.Application');book=None;record={};stop=threading.Event();dialogs=[]
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(p),UpdateLinks=0)
  configure(book,json.loads((REPO/'validation/G1R12_native25_20261009/inputs'/(case+'.json')).read_text(encoding='utf8')))
  book.Worksheets('要素定義').Range('A3').Value2=case+'・G1R17・TARGET200、q8、適応OFF。ゼロ応力頂点ファンの局所修復。'
  pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
  def monitor():
   while not stop.wait(.5):
    def window(h,_):
     if win32process.GetWindowThreadProcessId(h)[1]!=pid or not win32gui.IsWindowVisible(h) or win32gui.GetClassName(h)!='#32770':return
     text=[];buttons=[]
     def child(c,_):
      s=win32gui.GetWindowText(c)
      if s:text.append(s)
      if s=='OK':buttons.append(c)
     win32gui.EnumChildWindows(h,child,None)
     if buttons:
      dialogs.append(text);write(folder/'dialogs.json',dialogs)
      for c in buttons:win32gui.PostMessage(c,win32con.BM_CLICK,0,0)
    win32gui.EnumWindows(window,None)
  threading.Thread(target=monitor,daemon=True).start()
  ws=book.Worksheets('設定');assert ws.Range('B5').Value2==200 and ws.Range('B7').Value2==0
  assert book.Worksheets('浸透入力').Range('D4').Value2==0
  call=lambda n,*a:app.Run("'"+book.Name+"'!"+n,*a);call('RPX_NativeSetMode','auto')
  record={'state':'running','case':case,'excel_pid':pid,'source_sha256':sha(p),'DLL_SHA256':sha(folder/'rpfem_native_x64.dll'),'scope':case+' coarse TARGET200 q8 ADAPT0 only; no full run','native_start':str(call('RPX_NativeInfo'))};write(folder/'status.json',record);print('RUNNING COARSE',case,'PID',pid,flush=True)
  t=time.perf_counter();call('RPX_Run');record['seconds']=time.perf_counter()-t
  record.update(lower=book.Worksheets('解析結果').Range('B4').Value2,upper=book.Worksheets('解析結果').Range('C4').Value2,terminal_text=book.Worksheets('解析結果').Range('A1').Value2,native_end=str(call('RPX_NativeInfo')),dialogs=dialogs)
  assert isinstance(record['lower'],(float,int)) and 0<record['lower']<=record['upper'],record
  record['state']='Fs_obtained';book.Save();record['result_sha256']=sha(p);write(folder/'status.json',record);print(json.dumps(record,ensure_ascii=False),flush=True)
 except Exception:record['state']='failed';record['exception']=traceback.format_exc();raise
 finally:
  stop.set();write(folder/'status.json',record)
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
if __name__=='__main__':main()
