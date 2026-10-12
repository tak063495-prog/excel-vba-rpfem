"""Serial final fixed/whole/coarse verification; deliberately no full P06 batch."""
import os,sys,json,time,subprocess,traceback
from prepare import ROOT,REPO,sha,write

def run():
 core=json.loads((ROOT/'results/core_pipeline.json').read_text(encoding='utf8'))
 assert core['state']=='completed' and len(core['runs'])==22
 jobs=[['verify_fixed.py'],['run_levee.py','--backend','dll5','--tag','levee_dll5_final'],
       ['run_case.py','D07A','--backend','dll5','--tag','d07a_roots_dll5'],
       ['run_case.py','D03N','--backend','dll5','--tag','d03n_roots_dll5'],
       ['run_case.py','P06A','--backend','dll5','--tag','p06a_roots_dll5','--minimum-test-fs','0.125'],
       ['run_case.py','P06N','--backend','dll5','--tag','p06n_roots_dll5','--record-lower-bracket-failure'],
       ['verify_coarse.py'],['verify_release.py']]
 status=ROOT/'results/pipeline.json'
 record={'state':'running','dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'runs':[],
         'scope':'Fixed q8, fresh full levee, four coarse comparisons only. Full P06, tiny-Fs exploration and interrupted 26-case batch excluded.'}
 if '--resume' in sys.argv and status.exists():
  previous=json.loads(status.read_text(encoding='utf8'));assert previous['dll_sha256']==record['dll_sha256']
  record['runs']=[r for r in previous['runs'] if r['exit_code']==0]
  record['previous_failures']=previous.get('previous_failures',[])+[r for r in previous['runs'] if r['exit_code']!=0]
 try:
  for index,args in enumerate(jobs):
   current=' '.join(args)
   if any(r['command']==current for r in record['runs']):print('RETAIN PASS',current,flush=True);continue
   record['current']=current;write(status,record);print('START',current,flush=True)
   logfile=ROOT/'results'/('pipeline_'+str(index)+'.log')
   if logfile.exists():logfile.rename(logfile.with_name(logfile.stem+'_previous_'+str(time.time_ns())+'.log'))
   t=time.perf_counter()
   with logfile.open('w',encoding='utf8') as out:
    r=subprocess.run([sys.executable,str(ROOT/'tools'/args[0]),*args[1:]],cwd=REPO,stdout=out,stderr=subprocess.STDOUT,env={**os.environ,'PYTHONIOENCODING':'utf-8'})
   record['runs'].append({'command':current,'exit_code':r.returncode,'seconds':time.perf_counter()-t,'log':str(logfile.relative_to(ROOT))})
   write(status,record);print('END',current,r.returncode,flush=True)
   assert r.returncode==0,current+' failed; inspect '+str(logfile)
   assert sha(ROOT/'delivery/rpfem_native_x64.dll')==record['dll_sha256'],'Final DLL must remain frozen'
  record['state']='completed'
 except Exception:record['state']='failed';record['exception']=traceback.format_exc();raise
 finally:write(status,record)
if __name__=='__main__':run()
