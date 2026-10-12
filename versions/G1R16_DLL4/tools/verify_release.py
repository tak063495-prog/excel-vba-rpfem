"""Verify actual saved workbooks/logs and delivery source, then select final book.
Never rewrites XLSM internals: delivery is a byte copy of native Excel's saved result.
"""
import gc,json,shutil,struct,subprocess
import win32com.client
from prepare import ROOT,REPO,DLL3,write,sha
from verify_saved import verify_book,saved_cells,log,parts,compare_logs,macros,canonical
from native_helpers import compile_book

def duration(rows):return float(next(r['elapsed_s'] for r in reversed(rows) if r['event']=='run_end'))
def runlog(folder):
 paths=list((folder/'RPFEM_logs').glob('*analysis.tsv'));assert len(paths)==1;return paths[0]
def fs_exact(a,b):return all(struct.pack('<d',a[p])==struct.pack('<d',b[p]) for p in ['B4','C4'])
def result_cells_exact(a,b):
 return all((struct.pack('<d',a[p])==struct.pack('<d',b[p])) if isinstance(a.get(p),(int,float)) and isinstance(b.get(p),(int,float)) else a.get(p)==b.get(p) for p in ['B4','C4'])
def outcome(folder):
 s=json.loads((folder/'status.json').read_text(encoding='utf-8'));rows=log(runlog(folder));end=parts(next(r['detail'] for r in reversed(rows) if r['event']=='run_end'))
 origins=[parts(r['detail']) for r in rows if r['event']=='error_origin']
 return {'raw_harness_status':s['status'],'terminal_status':end['status'],'target_met':end['status']=='OPTIMAL','Fs_obtained':end['status'] in ['OPTIMAL','PARTIAL_CERTIFIED'] and isinstance(s.get('lower'),(int,float)) and s['lower']>0 and isinstance(s.get('upper'),(int,float)) and s['upper']>=s['lower'],'lower':s.get('lower'),'upper':s.get('upper'),'seconds':s.get('analysis_seconds'),'target':s.get('target'),'q':s.get('q'),'adapt':s.get('adapt'),'source_sha256':s['source_sha256'],'dll_sha256':s['dll_sha256'],'terminal_text':s.get('terminal_text'),'error_origins':origins,'upper_bracket':[parts(r['detail']) for r in rows if r['event']=='p4_final_bracket' and parts(r['detail']).get('mode')=='upper']}

def main():
 assert json.loads((ROOT/'results/pipeline.json').read_text(encoding='utf-8'))['state']=='completed'
 folder=ROOT/'results/levee_dll4_final';oldfolder=ROOT/'results/levee_dll3_serial'
 s=json.loads((folder/'status.json').read_text(encoding='utf-8'));oldstatus=json.loads((oldfolder/'status.json').read_text(encoding='utf-8'))
 assert s['status']==oldstatus['status']=='completed' and s['dll_sha256']==sha(ROOT/'delivery/rpfem_native_x64.dll')
 final=folder/'DLL4_Levee_Result.xlsm';previous=oldfolder/'DLL3_Levee_Result.xlsm'
 verification=verify_book(final)
 before=saved_cells(previous);after=saved_cells(final);reference=saved_cells(DLL3/'delivery/RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm')
 exact={n:before[n]==after[n]==reference[n] for n in ['要素定義','材料データ','設定','浸透入力','節点データ','要素データ']};assert all(exact.values()),exact
 assert fs_exact(before['解析結果'],after['解析結果']) and fs_exact(reference['解析結果'],after['解析結果'])
 comparison=compare_logs(runlog(oldfolder),runlog(folder));assert comparison['matches'],comparison['errors']
 vba=DLL3/'results/reference/vba_analysis.tsv';vbacomparison=compare_logs(vba,runlog(folder));assert vbacomparison['matches'],vbacomparison['errors']
 current=log(runlog(folder));old=log(runlog(oldfolder));seconds=duration(current);oldseconds=duration(old)
 phases=lambda rows:{parts(r['detail'])['phase']:float(parts(r['detail'])['seconds']) for r in rows if r['event']=='run_phase'}
 result={'DLL4_seconds':seconds,'DLL3_seconds':oldseconds,'DLL4_COM_seconds':s['analysis_seconds'],'DLL3_COM_seconds':oldstatus['analysis_seconds'],'reduction_percent_vs_DLL3':100*(1-seconds/oldseconds),'speedup_vs_DLL3':oldseconds/seconds,'VBA_existing_seconds':duration(log(vba)),'VBA_time_source':'Actual existing VBA log; no new full VBA speed run','Fs_bitwise_exact':True,'lower':s['lower'],'upper':s['upper'],'same_saved_inputs_and_mesh':exact,'DLL3_log_comparison':comparison,'VBA_log_comparison':vbacomparison,'DLL3_phases':phases(old),'DLL4_phases':phases(current),'dll_sha256':s['dll_sha256'],'timing_scope':'serial fresh processes, same PC, one pair; RPX_Run log; excludes separate seep/mesh calls, save and Python audit','seep_seconds':{'dll3':oldstatus['seep_seconds'],'dll4':s['seep_seconds']},'mesh_seconds':{'dll3':oldstatus['mesh_seconds'],'dll4':s['mesh_seconds']}}
 write(ROOT/'results/comparison.json',result)
 roots=[]
 for case in ['P06A','P06N','D03N','D07A']:
  a=ROOT/'results'/(case.lower()+'_roots_dll3');b=ROOT/'results'/(case.lower()+'_roots_dll4');oa=outcome(a);ob=outcome(b)
  if case=='P06A':
   assert oa['terminal_status']=='NUMERICAL_FAILURE' and oa['terminal_text'].endswith('HSD_NOT_CONVERGED')
   assert ob['terminal_status']=='CANCELLED' and all(r.get('number')=='18' and r.get('message')=='TEST_SCOPE_MIN_FS_FLOOR' for r in ob['error_origins'])
   la=compare_logs(runlog(a),runlog(b),15);assert la['matches'] and la['solves']==la['completed_solves']==15,la
   for folder2 in [a,b]:
    starts=[r for r in log(runlog(folder2)) if r['event']=='solve_start' and int(r['solve_id'])<=15];assert len(starts)==15 and all(float(r['strength_factor'])>=.125 for r in starts)
   starts_new=[r for r in log(runlog(b)) if r['event']=='solve_start'];assert len(starts_new)==15
   src=macros(ROOT/'delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm');probe=macros(b/'P06A_dll4.xlsm');assert set(src)==set(probe)
   assert [n for n in src if canonical(src[n])!=canonical(probe[n])]==['RPX_Fs']
   anchor='Private Function FsValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double) As Double'
   hook='\n    If mode = "lower" And factor < 0.125# Then\n        XCancel = True\n        Err.Raise 18, "DLL4TestScope", "TEST_SCOPE_MIN_FS_FLOOR"\n    End If'
   assert canonical(src['RPX_Fs'].replace(anchor,anchor+hook))==canonical(probe['RPX_Fs'])
   times={backend:float(next(r['elapsed_s'] for r in log(runlog(folder2)) if r['event']=='solve_end' and r['solve_id']=='15')) for backend,folder2 in [('DLL3',a),('DLL4',b)]}
   ca=saved_cells(a/'P06A_dll3.xlsm');cb=saved_cells(b/'P06A_dll4.xlsm');same=all(ca[n]==cb[n] for n in ['要素定義','材料データ','設定','浸透入力','節点データ','要素データ']);assert same
   roots.append({'case':case,'DLL3':oa,'DLL4':ob,'Fs_bitwise_exact':None,'comparison_passed':True,'case_passed':False,'full_case_compared':False,'comparison_scope':'15 complete solves with Fs>=0.125; original DLL3 failed later at solve23; DLL4 test-copy-only stop before solve16','log_comparison':la,'same_inputs_mesh':same,'comparison_seconds':times,'test_only_Fs_control_change_verified':True,'reduction_percent_vs_DLL3':100*(1-times['DLL4']/times['DLL3'])})
   continue
  assert oa['terminal_status']==ob['terminal_status'] and oa['Fs_obtained']==ob['Fs_obtained']
  if not oa['Fs_obtained']:
   assert oa['error_origins']==ob['error_origins'] and oa['error_origins']
   assert all(r.get('number')=='5' and r.get('message','').startswith('FS_LOWER_BRACKET_NOT_FOUND') for r in oa['error_origins'])
   assert oa['upper_bracket']==ob['upper_bracket']
  la=compare_logs(runlog(a),runlog(b));assert la['matches'],la['errors']
  ca=saved_cells(a/(case+'_dll3.xlsm'));cb=saved_cells(b/(case+'_dll4.xlsm'));assert result_cells_exact(ca['解析結果'],cb['解析結果'])
  roots.append({'case':case,'DLL3':oa,'DLL4':ob,'Fs_bitwise_exact':True if oa['Fs_obtained'] else None,'saved_result_cells_match':True,'comparison_passed':True,'case_passed':oa['Fs_obtained'],'full_case_compared':True,'comparison_scope':'Full original RPX_Run to identical terminal outcome','log_comparison':la,'same_inputs_mesh':all(ca[n]==cb[n] for n in ['要素定義','材料データ','設定','浸透入力','節点データ','要素データ']),'comparison_seconds':{'DLL3':oa['seconds'],'DLL4':ob['seconds']},'reduction_percent_vs_DLL3':100*(1-ob['seconds']/oa['seconds'])});assert roots[-1]['same_inputs_mesh']
 catalog={}
 for case in ['P06A','P06N','D03N','D07A']:
  name='validation/G1R12_native25_20261009/inputs/'+case+'.json';data=(REPO/name).read_bytes();assert subprocess.check_output(['git','show','HEAD:'+name],cwd=REPO)==data
  assert json.loads(data)['settings']['GRAVITY']=='固定';catalog[case]={'sha256':sha(REPO/name),'unchanged_from_git_baseline':True,'GRAVITY':'固定'}
 write(ROOT/'results/catalog_verification.json',catalog)
 p06=outcome(ROOT/'results/p06a_full_dll4');cancel=json.loads((ROOT/'results/p06a_full_dll4/cancellation.json').read_text(encoding='utf-8'));assert cancel['user_requested_stop'] and cancel['stop_macro_returned'] and p06['terminal_status']=='CANCELLED'
 write(ROOT/'results/root_regressions.json',{'paired_roots':roots,'P06A_full_mesh':p06,'P06A_full_mesh_note':'User explicitly requested cancellation after approximately 66 minutes because full runs can take 10 hours. RPX_Stop was used; not a solver failure, not a completion/accuracy/target pass. No full-size adaptive speed comparator. Final DLL with a pre-mode-switch-fix workbook; numerical code is the same. P06N full adaptive run not started.'})
 shutil.copyfile(final,ROOT/'delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm')
 app=win32com.client.DispatchEx('Excel.Application');book=None
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(ROOT/'delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm'),UpdateLinks=0)
  compile_book(app,book);assert book.VBProject.VBComponents.Count==69
  actions=[str(s.OnAction) for s in book.Worksheets('要素定義').Shapes if str(s.OnAction)]
  assert set(actions)=={'RPX_CreateMesh','RPX_Run','RPX_ShowDiagram','RPX_Stop'},actions
  verification['main_buttons_resolve_in_renamed_delivery_book']=actions
  msg=str(app.Run("'"+book.Name+"'!RPX_NativeInfo"));assert msg.startswith('DLL_READY');verification['reopen_native_info']=msg;verification['reopen_whole_project_compile']='PASS';verification['actual_module_count']=69;verification['security_unmodified']=app.AutomationSecurity
 finally:
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
 verification.update(verify_book(ROOT/'delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm'));write(ROOT/'results/delivery_verification.json',verification)
 print('RELEASE PASS',seconds,'DLL3',oldseconds,'reduction%',result['reduction_percent_vs_DLL3'],flush=True)
if __name__=='__main__':main()
