"""Finish a bounded P06A prefix comparison and the two remaining case pairs."""
import json,os,shutil,subprocess,sys,time
from prepare import ROOT,REPO,write
from verify_saved import log,parts
def main():
 p=ROOT/'results/pipeline.json';record=json.loads(p.read_text(encoding='utf-8'));assert record['state']=='failed'
 assert record['current'].startswith('run_case.py P06A --backend dll3 ')
 folder=ROOT/'results/p06a_roots_dll3';rows=log(next((folder/'RPFEM_logs').glob('*analysis.tsv')))
 end=parts(next(r['detail'] for r in reversed(rows) if r['event']=='run_end'))
 assert end['status']=='NUMERICAL_FAILURE' and end['message']=='HSD_NOT_CONVERGED'
 archive=ROOT/'results/pipeline_second_failure.json';assert not archive.exists();shutil.copyfile(p,archive)
 record['P06A_scope_resolution']='Optional additional coarse-mesh validation bounded after the scope question and stated recommended assumption. DLL3 ended at HSD_NOT_CONVERGED before stop; preserve raw failure. Compare all completed Fs>=0.125 points with DLL4; do not claim a full-P06A failure-path match or completed Fs.'
 jobs=[['run_case.py','P06A','--backend','dll4','--tag','p06a_roots_dll4','--target','200','--q','8','--adapt','0','--minimum-test-fs','0.125']]
 for case in ['D03N','D07A']:
  for backend in ['dll3','dll4']:
   jobs.append(['run_case.py',case,'--backend',backend,'--tag',case.lower()+'_roots_'+backend,'--target','200','--q','8','--adapt','0','--record-lower-bracket-failure'])
 for args in jobs:
  record['state']='running';record['current']=' '.join(args);write(p,record);t=time.perf_counter();print('START',record['current'],flush=True)
  path=ROOT/'results'/('pipeline_bounded_'+str(len(record['runs']))+'.log')
  with path.open('w',encoding='utf-8') as f:
   r=subprocess.run([sys.executable,str(ROOT/'tools'/args[0]),*args[1:]],cwd=REPO,stdout=f,stderr=subprocess.STDOUT,env={**os.environ,'PYTHONIOENCODING':'utf-8'})
  tag=args[args.index('--tag')+1];s=json.loads((ROOT/'results'/tag/'status.json').read_text(encoding='utf-8'))
  record['runs'].append({'command':' '.join(args),'exit_code':r.returncode,'wall_seconds':time.perf_counter()-t,'log':str(path.relative_to(ROOT)),'case_status':s['status'],'Fs_obtained':s['Fs_obtained'],'comparison_driver_outcome':s['comparison_driver_outcome']});write(p,record)
  if r.returncode:
   record['state']='failed';write(p,record);raise RuntimeError('Validation execution failed: '+tag)
  if args[1]=='P06A':write(ROOT/'results/p06a_root_pipeline.json',{'state':'bounded_prefix_completed','runs':[r for r in record['runs'] if r['command'].startswith('run_case.py P06A ')],'minimum_compared_fs':0.125,'full_case_compared':False})
  print('END',tag,s['status'],s['comparison_driver_outcome'],flush=True)
 record['state']='completed';write(p,record)
if __name__=='__main__':main()
