"""Resume after an observed baseline coarse-mesh Fs failure; preserve raw failure.
Only the named original bracket error can be compared as a failed case.
Cancellation, OOM, audit/subscript/unknown errors still abort validation.
"""
import json,os,shutil,subprocess,sys,time
from prepare import ROOT,REPO,write
from verify_saved import log,parts

def main():
 initial=ROOT/'results/pipeline.json';record=json.loads(initial.read_text(encoding='utf-8'))
 assert record['state']=='failed'
 assert record['current'].startswith('run_case.py P06N --backend dll3 ')
 folder=ROOT/'results/p06n_roots_dll3'
 status=json.loads((folder/'status.json').read_text(encoding='utf-8'))
 origins=[parts(r['detail']) for p in (folder/'RPFEM_logs').glob('*analysis.tsv') for r in log(p) if r['event']=='error_origin']
 assert origins and all(r.get('number')=='5' and r.get('message','').startswith('FS_LOWER_BRACKET_NOT_FOUND') for r in origins),origins
 assert status['status']=='failed'
 archive=ROOT/'results/pipeline_initial_failure.json';assert not archive.exists();shutil.copyfile(initial,archive)
 record.pop('exception',None);record['protocol_discrepancy']='The initial controller required positive Fs on every coarse mesh. Actual DLL3 cannot form the lower bracket for P06N TARGET200. Preserve its failed case, compare DLL4 under identical settings, and never count this as an Fs pass.'
 record['runs'][-1]['case_failed']=True;record['runs'][-1]['failure_classification']='BASELINE_FS_LOWER_BRACKET_NOT_FOUND'
 jobs=[('P06N','dll4')]+[(case,backend) for case in ['P06A','D03N','D07A'] for backend in ['dll3','dll4']]
 for case,backend in jobs:
  tag=case.lower()+'_roots_'+backend;args=['run_case.py',case,'--backend',backend,'--tag',tag,'--target','200','--q','8','--adapt','0','--record-lower-bracket-failure']
  record['state']='running';record['current']=' '.join(args);write(initial,record);t=time.perf_counter();print('START',record['current'],flush=True)
  path=ROOT/'results'/('pipeline_resumed_'+str(len(record['runs']))+'.log')
  with path.open('w',encoding='utf-8') as f:
   run=subprocess.run([sys.executable,str(ROOT/'tools'/args[0]),*args[1:]],cwd=REPO,stdout=f,stderr=subprocess.STDOUT,env={**os.environ,'PYTHONIOENCODING':'utf-8'})
  s=json.loads((ROOT/'results'/tag/'status.json').read_text(encoding='utf-8'))
  record['runs'].append({'command':' '.join(args),'exit_code':run.returncode,'wall_seconds':time.perf_counter()-t,'log':str(path.relative_to(ROOT)),'case_status':s['status'],'Fs_obtained':s['Fs_obtained'],'comparison_driver_outcome':s['comparison_driver_outcome']});write(initial,record)
  if run.returncode:
   record['state']='failed';write(initial,record);raise RuntimeError('Comparison execution failed; inspect '+tag)
  if case=='P06A':
   write(ROOT/'results/p06a_root_pipeline.json',{'state':'completed' if backend=='dll4' else 'running','runs':[r for r in record['runs'] if r['command'].startswith('run_case.py P06A ')]})
  print('END',case,backend,s['status'],s['comparison_driver_outcome'],flush=True)
 record['state']='completed';write(initial,record)

if __name__=='__main__':main()
