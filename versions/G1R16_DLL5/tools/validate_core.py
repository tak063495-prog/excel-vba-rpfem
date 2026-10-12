"""Final serial acceptance. Does not start any full P06 or 25/26-case batch.
Native Excel is serialized, including workbook close, before the next process.
"""
import os,sys,subprocess,time,traceback,json
from prepare import ROOT,REPO,write,sha

def run():
 jobs=[['test_guard_oracle.py'],['test_hsd_regression.py'],['test_bulk_oracle.py'],['test_nt_oracle.py'],['probe_cw.py'],['test_nt_oracle.py','--extended'],['test_lu_oracle.py'],['test_boundary.py'],['test_issue_regression.py'],['test_job_contract.py'],['test_new_job_contract.py'],['test_fallback_seep.py'],['test_mode_switch.py','--tag','mode_switch_final'],['test_job_lifecycle.py'],['test_native_faults.py'],['test_schur_python.py'],['test_krylov_python.py'],['test_hdd_python.py'],['test_compile_python.py'],['test_dd_core_python.py'],['test_new_job_lifecycle.py'],['benchmark_fixed.py','--tag','fixed_q8_release','--backends','dll5']]
 status=ROOT/'results/core_pipeline.json'
 record={'state':'running','dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'scope':'Final DLL5 targeted regressions; no full P06 restart; no interrupted 26-case batch restart','runs':[]}
 if '--resume' in sys.argv and status.exists():
  previous=json.loads(status.read_text(encoding='utf8'))
  assert previous['dll_sha256']==record['dll_sha256'],'A changed DLL requires a fresh acceptance run'
  record['runs']=previous['runs'];record['previous_failures']=previous.get('previous_failures',[])
  record['previous_failures'] += [r for r in record['runs'] if r['exit_code']!=0]
  record['runs']=[r for r in record['runs'] if r['exit_code']==0]
  record['resume_reason']='Resume the unchanged final DLL after correcting test-harness prerequisites/output isolation. Completed passes are retained; previous failures remain recorded.'
  record['resolved_harness_issues']=['NT initially compared Excel x87 0x137f with Python 0x027f. Aligning the Python test process gives zero mismatch; production floating-point mode unchanged.','New-job rerun correctly rejected existing output (OUTPUT_EXISTS); test output prefixes are now unique per run.']
 try:
  for index,args in enumerate(jobs):
   if any(r['command']==' '.join(args) for r in record['runs']):
    print('RETAIN PASS',' '.join(args),flush=True);continue
   record['current']=' '.join(args);write(ROOT/'results/core_pipeline.json',record);print('START',record['current'],flush=True)
   logfile=ROOT/'results'/('core_'+str(index)+'.log');t=time.perf_counter()
   if logfile.exists():
    backup=logfile.with_name(logfile.stem+'_previous_'+str(time.time_ns())+'.log')
    logfile.rename(backup)
    for previous in record.get('previous_failures',[]):
     if previous['log']==str(logfile.relative_to(ROOT)):
      previous['log']=str(backup.relative_to(ROOT))
   with logfile.open('w',encoding='utf8') as out:
    r=subprocess.run([sys.executable,str(ROOT/'tools'/args[0]),*args[1:]],cwd=REPO,stdout=out,stderr=subprocess.STDOUT,env={**os.environ,'PYTHONIOENCODING':'utf-8'})
   row={'command':record['current'],'exit_code':r.returncode,'seconds':time.perf_counter()-t,'log':str(logfile.relative_to(ROOT))};record['runs'].append(row);write(ROOT/'results/core_pipeline.json',record);print('END',record['current'],r.returncode,flush=True)
   assert r.returncode==0,record['current']+' failed; inspect '+str(logfile)
   assert record['dll_sha256']==sha(ROOT/'delivery/rpfem_native_x64.dll'),'DLL must remain frozen during acceptance'
  record['state']='completed'
 except Exception:record['state']='failed';record['exception']=traceback.format_exc();raise
 finally:write(ROOT/'results/core_pipeline.json',record)
if __name__=='__main__':run()
