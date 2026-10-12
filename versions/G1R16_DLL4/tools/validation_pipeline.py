"""Serial targeted acceptance and speed runs. No unrelated Excel is controlled.
Waits for the already-started full-size P06A test to finish before timing.
"""
import json,os,subprocess,sys,time,traceback
from prepare import ROOT,REPO,write

def run():
 record={'state':'waiting_for_full_P06A','runs':[]};write(ROOT/'results/pipeline.json',record)
 status=ROOT/'results/p06a_full_dll4/status.json'
 while json.loads(status.read_text(encoding='utf-8'))['status']=='running':time.sleep(10)
 jobs=[['benchmark_fixed.py','--tag','fixed_q8_threeway','--target','256','--q','8'],
       ['run_levee.py','--backend','dll3','--tag','levee_dll3_serial'],
       ['run_levee.py','--backend','dll4','--tag','levee_dll4_final']]
 for case in ['P06N','D03N','D07A']:
  for backend in ['dll3','dll4']:
   jobs.append(['run_case.py',case,'--backend',backend,'--tag',case.lower()+'_roots_'+backend,'--target','200','--q','8','--adapt','0'])
 try:
  for args in jobs:
   name=args[0]+' '+' '.join(args[1:]);record['state']='running';record['current']=name;write(ROOT/'results/pipeline.json',record);print('START',name,flush=True)
   logfile=ROOT/'results'/('pipeline_'+str(len(record['runs']))+'.log');t=time.perf_counter()
   with logfile.open('w',encoding='utf-8') as f:
    r=subprocess.run([sys.executable,str(ROOT/'tools'/args[0]),*args[1:]],cwd=REPO,stdout=f,stderr=subprocess.STDOUT,text=True,encoding='utf-8',env={**os.environ,'PYTHONIOENCODING':'utf-8'})
   row={'command':name,'exit_code':r.returncode,'wall_seconds':time.perf_counter()-t,'log':str(logfile.relative_to(ROOT))};record['runs'].append(row);write(ROOT/'results/pipeline.json',record);print('END',name,r.returncode,flush=True)
   if r.returncode:
    # A certified root interval that misses GAP is a distinct test outcome.
    if args[0]!='run_case.py':raise RuntimeError('Test failed: '+name)
    tag=args[args.index('--tag')+1];s=json.loads((ROOT/'results'/tag/'status.json').read_text(encoding='utf-8'))
    if not (isinstance(s.get('lower'),(int,float)) and isinstance(s.get('upper'),(int,float)) and s['lower']>0 and s['upper']>=s['lower']):raise RuntimeError('Fs search failed: '+name)
  record['state']='completed'
 except Exception:record['state']='failed';record['exception']=traceback.format_exc();raise
 finally:write(ROOT/'results/pipeline.json',record)
if __name__=='__main__':run()
