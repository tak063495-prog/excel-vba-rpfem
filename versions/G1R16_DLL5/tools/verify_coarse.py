"""Compare bounded/coarse searches with the immutable DLL4 evidence.
A matched unsuccessful Fs search is a regression comparison, not an Fs pass.
"""
import json,struct,subprocess,os
from pathlib import Path
from prepare import ROOT,REPO,DLL4,SOURCE,write,sha
from verify_saved import saved_cells,log,parts,compare_logs,macros,canonical
from verify_release import runlog

def read(path):return json.loads(path.read_text(encoding='utf8'))
def reference_cells(case,status):
 # Old test workbooks are intentionally not tracked/published. Preserve an
 # independently extracted, SHA-verified XML snapshot instead of repacking them.
 record=ROOT/'results/reference'/(case+'_DLL4_saved_cells.json')
 if not record.exists():
  location=Path(os.environ.get('RPFEM_DLL4_EVIDENCE_WORKTREE','C:/Users/link_/Downloads/RPFEM_G1R16_DLL4_20261012'))
  book=location/'versions/G1R16_DLL4/results'/(case.lower()+'_roots_dll4')/(case+'_dll4.xlsm')
  assert book.exists(),'Missing legacy workbook or published extracted snapshot: '+str(book)
  assert sha(book)==status['output_sha256'],'Legacy workbook must match its immutable published status'
  reference=macros(SOURCE);actual=macros(book);assert set(reference)==set(actual)
  differences=[n for n in reference if canonical(reference[n])!=canonical(actual[n])]
  if case=='P06A':
   assert differences==['RPX_Fs']
   anchor='Private Function FsValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double) As Double'
   hook='\n    If mode = "lower" And factor < 0.125# Then\n        XCancel = True\n        Err.Raise 18, "DLL4TestScope", "TEST_SCOPE_MIN_FS_FLOOR"\n    End If'
   assert canonical(reference['RPX_Fs'].replace(anchor,anchor+hook))==canonical(actual['RPX_Fs'])
  else:assert not differences,(case,differences)
  write(record,{'legacy_workbook_sha256':sha(book),'matches_published_DLL4_status':True,
   'source_matches_DLL4_except_verified_P06A_test_floor':True,'cells':saved_cells(book),
   'method':'Read-only XML extraction from actual native Excel saved test workbook; no workbook reconstruction.'})
 data=read(record);assert data['legacy_workbook_sha256']==status['output_sha256']
 assert data['matches_published_DLL4_status'] and data['source_matches_DLL4_except_verified_P06A_test_floor']
 return data['cells'],sha(record)
def main():
 records=[];catalog={}
 for case in ['D07A','D03N','P06A','P06N']:
  old=DLL4/'results'/(case.lower()+'_roots_dll4');new=ROOT/'results'/(case.lower()+'_roots_dll5')
  a=read(old/'status.json');b=read(new/'status.json')
  assert b['dll_sha256']==sha(ROOT/'delivery/rpfem_native_x64.dll') and a['dll_sha256']==sha(DLL4/'delivery/rpfem_native_x64.dll')
  assert a['target']==b['target']==200 and a['q']==b['q']==8 and a['adapt']==b['adapt']==0
  comp=compare_logs(runlog(old),runlog(new),15 if case=='P06A' else None);assert comp['matches'],(case,comp['errors'])
  end=lambda path:parts(next(r['detail'] for r in reversed(log(path)) if r['event']=='run_end'))['status']
  assert end(runlog(old))==end(runlog(new)),case
  ca,reference_sha=reference_cells(case,a);cb=saved_cells(new/(case+'_dll5.xlsm'))
  same={n:ca[n]==cb[n] for n in ['要素定義','材料データ','設定','浸透入力','節点データ','要素データ']};assert all(same.values()),(case,same)
  obtained=end(runlog(new)) in ['OPTIMAL','PARTIAL_CERTIFIED']
  if obtained:
   assert b['lower']>0 and b['upper']>=b['lower']
   assert all(struct.pack('<d',ca['解析結果'][p])==struct.pack('<d',cb['解析結果'][p]) for p in ['B4','C4'])
  old_time=a['analysis_seconds'];new_time=b['analysis_seconds']
  scope='Complete coarse search to the same terminal outcome; not full adaptive analysis.'
  if case=='P06A':
   starts=[r for r in log(runlog(new)) if r['event']=='solve_start']
   assert len(starts)==comp['completed_solves']==15 and all(float(r['strength_factor'])>=.125 for r in starts)
   assert b['comparison_driver_outcome']=='BOUNDED_TEST_STOP' and not obtained
   old_time=float(next(r['elapsed_s'] for r in log(runlog(old)) if r['event']=='solve_end' and r['solve_id']=='15'))
   new_time=float(next(r['elapsed_s'] for r in log(runlog(new)) if r['event']=='solve_end' and r['solve_id']=='15'))
   original=macros(ROOT/'delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm');probe=macros(new/(case+'_dll5.xlsm'))
   assert set(original)==set(probe) and [n for n in original if canonical(original[n])!=canonical(probe[n])]==['RPX_Fs']
   anchor='Private Function FsValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double) As Double'
   hook='\n    If mode = "lower" And factor < 0.125# Then\n        XCancel = True\n        Err.Raise 18, "DLL5TestScope", "TEST_SCOPE_MIN_FS_FLOOR"\n    End If'
   assert canonical(original['RPX_Fs'].replace(anchor,anchor+hook))==canonical(probe['RPX_Fs'])
   scope='Only the first 15 completed solves, Fs>=0.125. Test copy stops before solve16. No final Fs.'
  if case=='P06N':
   assert not obtained and b['comparison_driver_outcome']=='RECORDED_FAILED_FS_BRACKET'
   origins=lambda path:[parts(r['detail']) for r in log(path) if r['event']=='error_origin']
   assert origins(runlog(old))==origins(runlog(new))
   assert all(r['number']=='5' and r['message'].startswith('FS_LOWER_BRACKET_NOT_FOUND') for r in origins(runlog(new)))
  name='validation/G1R12_native25_20261009/inputs/'+case+'.json';data=(REPO/name).read_bytes()
  assert subprocess.check_output(['git','show','HEAD:'+name],cwd=REPO)==data
  catalog[case]={'sha256':sha(REPO/name),'unchanged_from_git_baseline':True}
  records.append({'case':case,'comparison_passed':True,'Fs_obtained':obtained,'target_met':end(runlog(new))=='OPTIMAL',
   'terminal_status':end(runlog(new)),'lower':b['lower'] if obtained else None,'upper':b['upper'] if obtained else None,
   'Fs_bitwise_exact':True if obtained else None,'scope':scope,'log_comparison':comp,'same_inputs_mesh':same,
   'seconds':{'DLL4_existing':old_time,'DLL5_current':new_time},'reduction_percent_vs_DLL4':100*(1-new_time/old_time),
   'DLL4_evidence':str(old.relative_to(REPO)),'DLL5_evidence':str(new.relative_to(ROOT)),
   'DLL4_native_saved_workbook_sha256':a['output_sha256'],'DLL4_extracted_cells_sha256':reference_sha,'DLL5_sha256':b['dll_sha256']})
 write(ROOT/'results/catalog_verification.json',catalog)
 write(ROOT/'results/root_regressions.json',{'cases':records,'P06A_full':'NOT_RUN: user stopped full analysis and tiny-Fs exploration.',
  'P06N_full':'NOT_RUN: no full adaptive P06 run restarted.','26_case_batch':'NOT_RUN: user interrupted this batch earlier.',
  'passed_comparisons':len(records),'scope':'Same coarse mesh and unchanged Fs policy; identical failures are recorded separately from successful Fs results.'})
 print('COARSE REGRESSION COMPARISONS PASS',len(records),flush=True)
if __name__=='__main__':main()
