"""Serial native Excel regression. NEVER starts full P06 or an Fs case batch."""
import subprocess,sys,time,traceback,json
from support import ROOT,REPO,write,sha
def main():
 record={'state':'running','DLL_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'runs':[],'scope':'Geometry25 only; controlled failures; two fixed-Fs VBA/DLL cases; analytic HSD/NT; no full analysis'}
 if '--resume'in sys.argv:
  old=json.loads((ROOT/'results/acceptance.json').read_text(encoding='utf8'));assert old['DLL_sha256']==record['DLL_sha256']
  record['runs']=[r for r in old['runs']if r['exit_code']==0];record['previous_failures']=old.get('previous_failures',[])+[r for r in old['runs']if r['exit_code']!=0]
  record['harness_correction']='Normalize interface encounter side labels for old/new mesh comparison. Actual 75 Excel exports are retained, never replaced by simulated results.'
 try:
  for name in ['test_geometry.py','test_guards.py','test_fixed_precision.py','test_mesh_fallback.py','test_hsd_regression.py','probe_cw.py','test_nt_oracle.py']:
   if any(r['test']==name for r in record['runs']):continue
   record['current']=name;write(ROOT/'results/acceptance.json',record);print('START',name,flush=True);t=time.perf_counter();p=ROOT/'results'/('acceptance_'+name+'.log')
   if p.exists():
    previous=p.with_name(p.stem+'_previous_'+str(time.time_ns())+'.log');p.rename(previous)
    for failure in record.get('previous_failures',[]):
     if failure['log']==str(p.relative_to(ROOT)):failure['log']=str(previous.relative_to(ROOT))
   args=['--verify-only']if name=='test_geometry.py'and '--resume'in sys.argv else []
   with p.open('w',encoding='utf8')as out:r=subprocess.run([sys.executable,str(ROOT/'tools'/name),*args],cwd=REPO,stdout=out,stderr=subprocess.STDOUT)
   record['runs'].append({'test':name,'exit_code':r.returncode,'seconds':time.perf_counter()-t,'log':str(p.relative_to(ROOT))});write(ROOT/'results/acceptance.json',record);print('END',name,r.returncode,flush=True)
   assert r.returncode==0,name+' failed; inspect '+str(p)
   assert sha(ROOT/'delivery/rpfem_native_x64.dll')==record['DLL_sha256']
  record['state']='PASS'
 except Exception:record['state']='failed';record['exception']=traceback.format_exc();raise
 finally:write(ROOT/'results/acceptance.json',record)
if __name__=='__main__':main()
