"""Independent audit of actual native saved upper/lower fields + root checks."""
import sys,json,numpy as np
from support import ROOT,write
sys.path.insert(0,str(ROOT/'tools'))
from inspect_saved import saved_cells,log,parts
from diagnose import mesh_from_arrays,solve
from common import physical_audit
def col(n):
 out=''
 while n:n,r=divmod(n-1,26);out=chr(65+r)+out
 return out
def saved_mesh_fields(path,rec):
 cells=saved_cells(path);nc=cells['節点データ'];ec=cells['要素データ']
 nr=sorted(int(k[1:])for k in nc if k.startswith('A')and k[1:].isdigit()and k!='A1');er=sorted(int(k[1:])for k in ec if k.startswith('A')and k[1:].isdigit()and k!='A1')
 xy=np.array([[nc['B'+str(i)],nc['C'+str(i)]]for i in nr]);tri=np.array([[int(ec[c+str(i)])-1 for c in ['B','C','D']]for i in er],np.int32);ids=np.array([int(ec['E'+str(i)])-1 for i in er],np.int32)
 m=mesh_from_arrays(xy,tri,ids,rec)
 upper=np.array([[ec[col(j)+str(i)]for j in range(8,20)]for i in er]).ravel()
 lower=np.array([[ec[col(j)+str(i)]for j in range(34,52)]for i in er]).ravel()/m.stress_scale
 return m,upper,lower
def main():
 out={'cases':[],'scope':'Saved native fields on actual coarse mesh. XML decimal serialization is audited, never used as an exact binary identity key.'}
 for case in ['P06N','P06A']:
  folder=ROOT/'results'/('coarse_g1r17_'+case);meta=json.loads((folder/'status.json').read_text(encoding='utf8'));assert meta['state']=='Fs_obtained'
  path=next(folder.glob('*.xlsm'));rec=json.loads((ROOT/'data/inputs'/(case+'.json')).read_text(encoding='utf8'));m,u,l=saved_mesh_fields(path,rec)
  rows=log(next((folder/'RPFEM_logs').glob('*.tsv')));ends={}
  for r in rows:
   if r['event']=='p4_final_bracket':ends[r['mode']]=parts(r['detail'])
  load=float(ends['lower']['target'])+float(ends['lower']['fa']);lower=np.r_[l,load]
  audits={'lower':physical_audit(m,lower,'lower',meta['lower'],8),'upper':physical_audit(m,u,'upper',meta['upper'],8)}
  assert all(a['pass']for a in audits.values()),(case,audits)
  entry={'case':case,'nodes':len(m.xy),'elements':len(m.tri),'Fs_lower':meta['lower'],'Fs_upper':meta['upper'],'independent_native_field_audits':audits,'native_brackets':ends,'python_points':[]};out['cases'].append(entry)
  for side in ['lower','upper']:
   for endpoint in ['a','b']:
    fs=float(ends[side][endpoint])
    # Fixed-lambda feasibility avoids overstating AlmostSolved optimality.
    r,x,p,z=solve(m,fs,side,fixed_load=1. if side=='lower'else None);r['endpoint']=endpoint;entry['python_points'].append(r)
    print(case,side,endpoint,fs,r['status'],r['objective'],r.get('physical_audit',{}).get('pass'),flush=True)
    if side=='lower'and endpoint=='a':assert r.get('physical_audit',{}).get('pass')and abs(r['load']-1)<1e-7
    if side=='lower'and endpoint=='b':
     assert r['status']in ['PrimalInfeasible','AlmostPrimalInfeasible'],r
     certificate=r['infeasibility_certificate'];strict=certificate['SOC_dual_error']<1e-7 and certificate['stationarity_max']<1e-7 and certificate['rhs_dot_dual']<-.99
     r['strict_infeasibility_certificate_pass']=strict
     if not strict:
      # Do not accept an inaccurate infeasibility certificate by its label.
      # Verify the optimized dual independently for the maximum-load bound.
      # A failed primal audit is kept as failed and its field is never accepted.
      check,unused_x,unused_p,unused_z=solve(m,fs,'lower')
      r['fallback_max_load_check']=check
      check['primal_field_accepted']=check.get('physical_audit',{}).get('pass',False)
      assert check['dual_stationarity_max']<1e-7 and check['dual_cone_error']<1e-7
      assert -check['solver_dual_obj']<1.-1e-5 and check['dual_gap']<1e-6
      check['dual_max_load_check_pass']=True
    if side=='upper':
     assert r.get('physical_audit',{}).get('pass'),r
     target=float(ends[side]['target']);assert (r['objective']>target if endpoint=='a'else r['objective']<target)
    write(ROOT/'results/coarse_python_verification.json',out)
 out['PASS']=True;write(ROOT/'results/coarse_python_verification.json',out)
 print('COARSE PYTHON FIELD/ROOT CHECKS PASS',flush=True)
if __name__=='__main__':main()
