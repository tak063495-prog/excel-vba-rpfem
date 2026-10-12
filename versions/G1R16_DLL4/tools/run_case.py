import argparse,gc,json,shutil,time,threading,traceback
from pathlib import Path
import win32com.client,win32process,win32gui,win32con
from prepare import ROOT,REPO,DLL3,write,sha
from test_hsd_regression import setvalue
from verify_saved import log,parts
from native_helpers import compile_book
def run():
 parser=argparse.ArgumentParser();parser.add_argument('case');parser.add_argument('--backend',choices=['dll3','dll4'],default='dll4');parser.add_argument('--tag',required=True);parser.add_argument('--target',type=int,default=1536);parser.add_argument('--q',type=int,default=8);parser.add_argument('--adapt',type=int,default=1);parser.add_argument('--record-lower-bracket-failure',action='store_true',help='Record the original FS_LOWER_BRACKET_NOT_FOUND as a failed case for a baseline comparison, never as an Fs pass.');parser.add_argument('--minimum-test-fs',type=float,default=0,help='Test-copy-only stop before a lower Fs. Original product source is unchanged.');args=parser.parse_args()
 assert args.case in ['P06A','P06N','D03N','D07A'] and args.tag.isidentifier()
 folder=ROOT/'results'/args.tag;folder.mkdir(parents=True,exist_ok=False)
 kit=DLL3 if args.backend=='dll3' else ROOT
 source=kit/'delivery'/('RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm' if args.backend=='dll3' else 'RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm')
 path=folder/(args.case+'_'+args.backend+'.xlsm');shutil.copyfile(source,path);shutil.copyfile(kit/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
 rec=json.loads((REPO/'validation/G1R12_native25_20261009/inputs'/(args.case+'.json')).read_text(encoding='utf-8'))
 assert rec['settings']['GRAVITY']=='固定'
 result={'case':args.case,'backend':args.backend,'status':'preparing','source_sha256':sha(source),'dll_sha256':sha(folder/'rpfem_native_x64.dll'),'catalog_sha256':sha(REPO/'validation/G1R12_native25_20261009/inputs'/(args.case+'.json')),'target':args.target,'q':args.q,'adapt':args.adapt,'scope':'full RPX_Run; original catalog geometry and materials; current G1R16 gravity fixed; seismic/seep off','gravity_note':'Original UTF-8 catalog GRAVITY is fixed; explicitly assigned identically on DLL3 and DLL4.'}
 app=win32com.client.DispatchEx('Excel.Application');book=None;stopped=threading.Event();dialogs=[]
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
  if args.minimum_test_fs:
   assert 0<args.minimum_test_fs<1
   module=book.VBProject.VBComponents('RPX_Fs').CodeModule;text=module.Lines(1,module.CountOfLines)
   anchor='Private Function FsValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double) As Double'
   assert text.count(anchor)==1
   hook='\r\n    If mode = "lower" And factor < '+format(args.minimum_test_fs,'.17g')+'# Then\r\n        XCancel = True\r\n        Err.Raise 18, "DLL4TestScope", "TEST_SCOPE_MIN_FS_FLOOR"\r\n    End If'
   module.DeleteLines(1,module.CountOfLines);module.AddFromString(text.replace(anchor,anchor+hook));compile_book(app,book)
   result['test_only_instrumentation']={'module':'RPX_Fs','minimum_fs':args.minimum_test_fs,'effect':'Raise cancellation18 before FsValue below floor; no numerical model, tolerance or product source change.'}
  pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1];result['excel_pid']=pid
  def monitor():
   while not stopped.wait(.5):
    def window(hwnd,_):
     if win32process.GetWindowThreadProcessId(hwnd)[1]!=pid or not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd)!='#32770':return
     texts=[];buttons=[]
     def child(h,_):
      value=win32gui.GetWindowText(h)
      if value:texts.append(value)
      if value in ['OK','はい(&Y)']:buttons.append(h)
     win32gui.EnumChildWindows(hwnd,child,None)
     if buttons:
      dialogs.append(texts);write(folder/'dialogs.json',dialogs)
      for button in buttons:win32gui.PostMessage(button,win32con.BM_CLICK,0,0)
    win32gui.EnumWindows(window,None)
  thread=threading.Thread(target=monitor,daemon=True);thread.start()
  app.EnableEvents=False
  ws=book.Worksheets('要素定義');ws.Range('A7:T1006').ClearContents()
  for first,last,key in [('A','C','points'),('E','H','regions'),('J','T','faces')]:
   ws.Range(first+'7:'+last+str(6+len(rec[key]))).Value2=tuple(tuple(x) for x in rec[key])
  ws=book.Worksheets('材料データ');ws.Range('A5:Q1004').ClearContents();ws.Range('A5:E'+str(4+len(rec['materials']))).Value2=tuple(tuple(x) for x in rec['materials'])
  settings=dict(rec['settings']);settings.update(GRAVITY='固定',TARGET=args.target,QUADRATURE=args.q,ADAPT=args.adapt,SEISMIC_KH=0,HSD_CAPTURE=0,HSD_SOLVE_SECONDS=7200)
  for key,value in settings.items():setvalue(book,key,value)
  book.Worksheets('浸透入力').Range('D4').Value2=0
  app.EnableEvents=True
  call=lambda name,*a:app.Run("'"+book.Name+"'!"+name,*a)
  call('RPX_NativeSetMode','auto');result['native_start']=str(call('RPX_NativeInfo'));assert result['native_start'].startswith('DLL_READY')
  book.Save();result['input_book_sha256']=sha(path);result['status']='running';write(folder/'status.json',result);print('RUNNING',args.case,args.backend,'PID',pid,flush=True)
  t=time.perf_counter();call('RPX_CreateMesh');result['mesh_seconds']=time.perf_counter()-t
  t=time.perf_counter();call('RPX_Run');result['analysis_seconds']=time.perf_counter()-t
  result['output']=book.Worksheets('解析結果').Range('A1:H40').Value2
  result['lower']=book.Worksheets('解析結果').Range('B4').Value2;result['upper']=book.Worksheets('解析結果').Range('C4').Value2;result['terminal_text']=book.Worksheets('解析結果').Range('A1').Value2
  rows=[r for p in (folder/'RPFEM_logs').glob('*analysis.tsv') for r in log(p)];endings=[r for r in rows if r['event']=='run_end']
  result['terminal_status']=parts(endings[-1]['detail'])['status']
  result['status']={'OPTIMAL':'completed','PARTIAL_CERTIFIED':'partial_certified'}.get(result['terminal_status'],'failed')
  result['target_met']=result['terminal_status']=='OPTIMAL'
  origins=[parts(r['detail']) for r in rows if r['event']=='error_origin'];result['error_origins']=origins
  known_bracket_failure=bool(origins) and all(r.get('number')=='5' and r.get('message','').startswith('FS_LOWER_BRACKET_NOT_FOUND') for r in origins)
  bounded_stop=bool(args.minimum_test_fs and origins) and all(r.get('number')=='18' and r.get('message')=='TEST_SCOPE_MIN_FS_FLOOR' for r in origins) and result['terminal_status']=='CANCELLED'
  result['Fs_obtained']=result['status'] in ['completed','partial_certified'] and isinstance(result['lower'],(int,float)) and result['lower']>0 and isinstance(result['upper'],(int,float)) and result['upper']>=result['lower']
  result['comparison_driver_outcome']='BOUNDED_TEST_STOP' if bounded_stop else 'RECORDED_FAILED_FS_BRACKET' if args.record_lower_bracket_failure and known_bracket_failure and result['status']=='failed' else result['status']
  result['native_end']=str(call('RPX_NativeInfo'));result['dialogs']=dialogs;book.Save();result['output_sha256']=sha(path)
  write(folder/'status.json',result);print(json.dumps({k:result[k] for k in ['case','backend','status','analysis_seconds','lower','upper','terminal_text']},ensure_ascii=False),flush=True)
  assert (result['status'] in ['completed','partial_certified'] and result['Fs_obtained']) or result['comparison_driver_outcome'] in ['RECORDED_FAILED_FS_BRACKET','BOUNDED_TEST_STOP'],result['terminal_text']
 except Exception:
  result['exception']=traceback.format_exc();result['status']='failed';raise
 finally:
  stopped.set();write(folder/'status.json',result)
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
if __name__=='__main__':run()
