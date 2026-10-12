"""Serial P06A coarse-mesh roots after the primary comparison pipeline."""
import json,os,subprocess,sys,time
from prepare import ROOT,REPO,write
def main():
 record={'state':'waiting','runs':[]};write(ROOT/'results/p06a_root_pipeline.json',record)
 while True:
  p=json.loads((ROOT/'results/pipeline.json').read_text(encoding='utf-8'))
  if p['state']=='failed':raise RuntimeError('Primary comparison pipeline failed; inspect evidence before continuing')
  if p['state']=='completed':break
  time.sleep(5)
 for backend in ['dll3','dll4']:
  record['state']='running';record['backend']=backend;write(ROOT/'results/p06a_root_pipeline.json',record)
  command=[sys.executable,str(ROOT/'tools/run_case.py'),'P06A','--backend',backend,'--tag','p06a_roots_'+backend,'--target','200','--q','8','--adapt','0']
  with (ROOT/'results'/('p06a_root_'+backend+'.log')).open('w',encoding='utf-8') as f:
   r=subprocess.run(command,cwd=REPO,stdout=f,stderr=subprocess.STDOUT,env={**os.environ,'PYTHONIOENCODING':'utf-8'})
  record['runs'].append({'backend':backend,'exit_code':r.returncode});write(ROOT/'results/p06a_root_pipeline.json',record);assert r.returncode==0
 record['state']='completed';write(ROOT/'results/p06a_root_pipeline.json',record);print('P06A root pair completed',flush=True)
if __name__=='__main__':main()
