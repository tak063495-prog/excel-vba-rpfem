import gc,json,shutil
import win32com.client
from prepare import ROOT,write,sha
from test_hsd_regression import inject,setvalue
from native_helpers import compile_book
def run():
 folder=ROOT/'results/job_contract';folder.mkdir(parents=True,exist_ok=True)
 path=folder/'Job_Contract.xlsm';shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm',path);shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
 app=win32com.client.DispatchEx('Excel.Application');book=None;rows=[]
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0);inject(book,'JobContract')
  for name,marker in [('RPX_LDL','    RPX_NativeCheckCode NRBegin(refineHandle, orderN, refineEpoch, VarPtr(rhs(0)))'),('RPX_Solve','    serial = RPX_NativeHLinearBegin(rhs, opArgs, kind, hFactorEpoch, hUnshifted)')]:
   cm=book.VBProject.VBComponents(name).CodeModule;s=cm.Lines(1,cm.CountOfLines).replace('\r\n','\n');assert s.count(marker)==1;s=s.replace(marker,marker+'\n    DLL4Inject',1);cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(s.replace('\n','\r\n'))
  ws=book.Worksheets('設定')
  for r in range(4,101):
   if ws.Cells(r,4).Value2 is None and ws.Cells(r,2).Value2 is None:
    ws.Cells(r,4).Value2='HSD_STAGE_SECONDS';ws.Cells(r,2).Value2=0;break
  else:raise ValueError('No empty setting row for test')
  setvalue(book,'DIAGNOSTIC_MODE',1);compile_book(app,book)
  call=lambda name,*args:app.Run("'"+book.Name+"'!"+name,*args)
  original=book.Worksheets('材料データ').Cells(5,2).Value2
  for hsd in [False,True]:
   for fault in ['cancel','memory','subscript','unknown','input','']:
    app.EnableEvents=False;book.Worksheets('材料データ').Cells(5,2).Value2=original
    msg=str(call('JobCase',hsd,fault,str(folder/('hsd' if hsd else 'normal'))));print(hsd,fault or 'normal',msg,flush=True)
    expected={'cancel':'ERROR 18:ANALYSIS_CANCELLED','memory':'ERROR 7:INJECTED_MEMORY','subscript':'ERROR 9:INJECTED_SUBSCRIPT','unknown':'ERROR 513:INJECTED_UNKNOWN','input':'ERROR 5:INPUT_MISMATCH','':'PASS'}[fault]
    row={'hsd':hsd,'fault':fault,'result':msg,'pass':msg.startswith(expected) and 'h=0;r=0;l=0' in msg};rows.append(row);write(folder/'summary.json',{'tests':rows,'dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll')});assert row['pass'],row
    if fault=='input':assert book.Worksheets('材料データ').Cells(5,2).Value2==12345.,'input must not be restored by solver'
  app.EnableEvents=False;book.Worksheets('材料データ').Cells(5,2).Value2=original
 finally:
  write(folder/'summary.json',{'tests':rows,'dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'passed':sum(x['pass'] for x in rows)})
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
 print('Job contracts PASS',len(rows),flush=True)
if __name__=='__main__':run()
