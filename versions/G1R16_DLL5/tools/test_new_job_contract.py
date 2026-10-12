"""Actual Excel aborts after native compiler/Schur/Krylov context creation.
Includes events-off sheet mutation; no production test hooks are delivered.
"""
import gc,json,shutil,time
import win32com.client
from prepare import ROOT,REPO,write,sha
from native_helpers import compile_book
from test_hsd_regression import inject,setvalue
from benchmark_fixed import configure

def run():
 folder=ROOT/'results/new_job_contract';folder.mkdir(exist_ok=True)
 outputs=folder/('outputs_'+str(time.time_ns()));outputs.mkdir()
 path=folder/'NewJobs.xlsm';shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm',path);shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
 app=win32com.client.DispatchEx('Excel.Application');book=None;records=[]
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
  for name in ['JobContract','HSDRegression','MeshProbe']:inject(book,name)
  hooks=[('RPX_Conic','    RPX_NativeCheckCode NCInit(XN, XM, XB, coneN, XEPtr(XM), VarPtr(args(0)), handle)','compile'),('RPX_Schur','    RPX_NativeCheckCode NSSolveBegin(nativeSchur, CLngLng(P4Pattern), CLngLng(factorEpoch), VarPtr(rhs(0)), Abs(CLng(checked)), Abs(CLng(P4ReuseU)), serial)','schur'),('RPX_Schur','    RPX_NativeCheckCode NSFactorBegin(nativeSchur, CLngLng(P4Pattern), CLngLng(P4Matrix))','schur_factor'),('RPX_Solve','    serial = RPX_NativeKrylovBegin(bdd, answer, bestError, hFactorEpoch, kind, hUnshifted)','krylov')]
  for name,marker,scope in hooks:
   cm=book.VBProject.VBComponents(name).CodeModule;s=cm.Lines(1,cm.CountOfLines).replace('\r\n','\n');assert s.count(marker)==1;s=s.replace(marker,marker+'\n    DLL5Inject "'+scope+'"');cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(s.replace('\n','\r\n'))
  cm=book.VBProject.VBComponents('HSDRegression').CodeModule;s=cm.Lines(1,cm.CountOfLines);marker='Call RPX_ReadInputs: Call RPX_GenerateMesh';assert s.count(marker)==1;s=s.replace(marker,marker+': Call RPX_InputBegin');cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(s)
  ws=book.Worksheets('設定')
  for r in range(4,101):
   if ws.Cells(r,4).Value2 is None and ws.Cells(r,2).Value2 is None:ws.Cells(r,4).Value2='HSD_STAGE_SECONDS';ws.Cells(r,2).Value2=0;break
  else:raise ValueError('Missing free setting row')
  compile_book(app,book);call=lambda name,*a:app.Run("'"+book.Name+"'!"+name,*a)
  zero='h=0;r=0;l=0;s=0;k=0;c=0'
  expected={'cancel':'ERROR 18:ANALYSIS_CANCELLED','memory':'ERROR 7:INJECTED_MEMORY','subscript':'ERROR 9:INJECTED_SUBSCRIPT','unknown':'ERROR 513:INJECTED_UNKNOWN','input':'ERROR 5:INPUT_MISMATCH','':'PASS'}
  original=book.Worksheets('材料データ').Cells(5,2).Value2
  for scope in ['compile','schur_factor','schur','krylov']:
   for fault in ['cancel','memory','subscript','unknown','input','']:
    call('RPX_InputEnd');app.EnableEvents=False;book.Worksheets('材料データ').Cells(5,2).Value2=original
    if scope=='compile':
     msg=str(call('JobCase',True,fault,str(outputs/'compile'),'compile'))
    else:
     case='P06A' if scope=='krylov' else 'P01A';rec=json.loads((REPO/'validation/G1R12_native25_20261009/inputs'/(case+'.json')).read_text(encoding='utf8'))
     if case=='P01A':rec['materials']=[[1,5.,45.,20.,45.]]
     configure(book,rec,64,2);original=book.Worksheets('材料データ').Cells(5,2).Value2
     setvalue(book,'DIAGNOSTIC_MODE',1);setvalue(book,'HSD_STAGE_SECONDS',1);call('RPX_NativeSetMode','auto');call('RPX_StageConfigure');call('DLL5Arm',scope,fault)
     app.EnableEvents=False
     msg=str(call('FixedFem','lower',1. if scope=='krylov' else 3.,str(outputs/(scope+'_'+(fault or 'ok')))))
     cached=str(call('DLL5Handles'))
     if fault=='':call('RPX_SchurRelease')
     msg+=';hits='+str(call('DLL5HitCount'))+';'+str(call('DLL5Handles'))
    ok=msg.startswith(expected[fault]) and zero in msg and 'hits=1;' in msg
    row={'scope':scope,'fault':fault,'message':msg,'pass':ok};records.append(row);write(folder/'summary.json',{'tests':records,'dll_sha256':sha(folder/'rpfem_native_x64.dll')});print(scope,fault or 'ok',msg,flush=True);assert ok,row
    if fault=='input':assert book.Worksheets('材料データ').Cells(5,2).Value2==12345.
  call('RPX_InputEnd')
 finally:
  write(folder/'summary.json',{'tests':records,'passed':sum(x['pass'] for x in records),'dll_sha256':sha(folder/'rpfem_native_x64.dll'),'test_only_native_context_abort_hooks':True})
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
 print('NEW JOB CONTRACT PASS',len(records),flush=True)
if __name__=='__main__':run()
