"""Validate independent references, then copy the native Excel saved result.
No spreadsheet writer or vbaProject.bin reconstruction is used.
"""
import gc,json,shutil,struct
import win32com.client
from prepare import ROOT,REPO,DLL3,DLL4,SOURCE,write,sha
from verify_saved import verify_book,saved_cells,log,parts,compare_logs
from native_helpers import compile_book

def read(path):return json.loads(path.read_text(encoding='utf8'))
def runlog(folder):
 paths=list((folder/'RPFEM_logs').glob('*analysis.tsv'));assert len(paths)==1
 return paths[0]
def duration(rows):return float(next(r['elapsed_s'] for r in reversed(rows) if r['event']=='run_end'))
def fs_exact(a,b):return all(struct.pack('<d',a[p])==struct.pack('<d',b[p]) for p in ['B4','C4'])
def phases(rows):return {parts(r['detail'])['phase']:float(parts(r['detail'])['seconds']) for r in rows if r['event']=='run_phase'}

def main():
 core=read(ROOT/'results/core_pipeline.json');assert core['state']=='completed' and len(core['runs'])==22
 assert all(r['exit_code']==0 for r in core['runs']) and core['dll_sha256']==sha(ROOT/'delivery/rpfem_native_x64.dll')
 fixed=read(ROOT/'results/fixed_comparison.json');assert fixed['passed'] and fixed['final_DLL5_sha256']==core['dll_sha256']
 oldfolder=ROOT/'results/levee_dll4_serial';folder=ROOT/'results/levee_dll5_final'
 oldstatus=read(oldfolder/'status.json');s=read(folder/'status.json')
 assert s['status']==oldstatus['status']=='completed' and s['dll_sha256']==core['dll_sha256']
 assert oldstatus['source_sha256']==sha(SOURCE) and oldstatus['dll_sha256']==sha(DLL4/'delivery/rpfem_native_x64.dll')
 final=folder/'DLL5_Levee_Result.xlsm';previous=oldfolder/'DLL4_Levee_Result.xlsm'
 verification=verify_book(final)
 before=saved_cells(previous);after=saved_cells(final);reference=saved_cells(SOURCE)
 names=['要素定義','材料データ','設定','浸透入力','節点データ','要素データ']
 exact={n:before[n]==after[n]==reference[n] for n in names};assert all(exact.values()),exact
 assert fs_exact(before['解析結果'],after['解析結果']) and fs_exact(reference['解析結果'],after['解析結果'])
 comparisons={'DLL4_fresh':compare_logs(runlog(oldfolder),runlog(folder)),
              'DLL3_existing':compare_logs(runlog(DLL4/'results/levee_dll3_serial'),runlog(folder)),
              'VBA_existing':compare_logs(DLL3/'results/reference/vba_analysis.tsv',runlog(folder))}
 assert all(c['matches'] for c in comparisons.values()),comparisons
 assert comparisons['DLL4_fresh']['solves']==comparisons['DLL4_fresh']['completed_solves']==53
 assert comparisons['DLL4_fresh']['audits']==603
 current=log(runlog(folder));old=log(runlog(oldfolder));seconds=duration(current);oldseconds=duration(old)
 times={'VBA_existing':duration(log(DLL3/'results/reference/vba_analysis.tsv')),
        'DLL3_existing':duration(log(runlog(DLL4/'results/levee_dll3_serial'))),
        'DLL4_fresh':oldseconds,'DLL5_fresh':seconds}
 result={'seconds':times,'DLL5_COM_seconds':s['analysis_seconds'],'DLL4_COM_seconds':oldstatus['analysis_seconds'],
  'reduction_percent_vs_DLL4':100*(1-seconds/oldseconds),'speedup_vs_DLL4':oldseconds/seconds,
  'reduction_percent_vs_VBA_reference':100*(1-seconds/times['VBA_existing']),
  'speedup_vs_VBA_reference':times['VBA_existing']/seconds,
  'VBA_time_source':'Existing actual log, not a new measurement; DLL3 also existing. DLL4 and DLL5 freshly measured in separate serialized Excel processes.',
  'Fs_bitwise_exact':True,'lower':s['lower'],'upper':s['upper'],'same_saved_inputs_and_mesh':exact,
  'log_comparisons':comparisons,'DLL4_phases':phases(old),'DLL5_phases':phases(current),
  'dll_sha256':s['dll_sha256'],'timing_scope':'One fresh serial pair, same PC; RPX_Run run_end. Excludes separate seep/mesh calls, save, Python audit.',
  'seep_seconds':{'dll4':oldstatus['seep_seconds'],'dll5':s['seep_seconds']},
  'mesh_seconds':{'dll4':oldstatus['mesh_seconds'],'dll5':s['mesh_seconds']}}
 write(ROOT/'results/comparison.json',result)
 destination=ROOT/'delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm'
 shutil.copyfile(final,destination);assert sha(destination)==sha(final)
 app=win32com.client.DispatchEx('Excel.Application');book=None
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(destination),UpdateLinks=0)
  compile_book(app,book);assert book.VBProject.VBComponents.Count==70
  actions=[str(s.OnAction) for s in book.Worksheets('要素定義').Shapes if str(s.OnAction)]
  assert set(actions)=={'RPX_CreateMesh','RPX_Run','RPX_ShowDiagram','RPX_Stop'},actions
  verification['main_buttons_resolve_in_renamed_delivery_book']=actions
  msg=str(app.Run("'"+book.Name+"'!RPX_NativeInfo"));assert msg.startswith('DLL_READY')
  verification.update(reopen_native_info=msg,reopen_whole_project_compile='PASS',actual_module_count=70,security_unmodified=app.AutomationSecurity)
 finally:
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
 assert sha(destination)==sha(final),'Read-only reopening must not resave the delivery book'
 verification.update(verify_book(destination));write(ROOT/'results/delivery_verification.json',verification)
 print('RELEASE PASS DLL5',seconds,'DLL4',oldseconds,'reduction%',result['reduction_percent_vs_DLL4'],flush=True)
if __name__=='__main__':main()
