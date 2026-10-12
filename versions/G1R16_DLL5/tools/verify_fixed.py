"""Require final DLL5 fields/mesh to match both actual VBA and DLL4 runs.
One-backend benchmark output alone is never treated as a comparison proof.
"""
import json
from prepare import ROOT,write,sha,DLL3

def run():
 initial=ROOT/'results/compiled_q8_first';final=ROOT/'results/fixed_q8_release'
 old=json.loads((initial/'summary.json').read_text(encoding='utf8'));new=json.loads((final/'summary.json').read_text(encoding='utf8'));records={}
 assert set(old['cases'])==set(new['cases']) and len(new['cases'])==6
 assert new['backends']['dll5']['dll_sha256']==sha(ROOT/'delivery/rpfem_native_x64.dll')
 for key,data in new['cases'].items():
  row=data['dll5'];assert row['independent_audit']['pass'] and row['Fs_input']==old['cases'][key]['vba']['Fs_input']
  for kind in ['field','mesh']:
   name=key+'_'+kind+'.bin';reference=(initial/'vba'/name).read_bytes()
   assert reference==(initial/'dll4'/name).read_bytes()==(final/'dll5'/name).read_bytes(),(key,kind)
  seconds={b:old['cases'][key][b]['seconds'] for b in ['vba','dll4']};seconds['dll5']=row['seconds']
  records[key]={'Fs_input':row['Fs_input'],'elements':row['elements'],'VBA_DLL4_DLL5_full_field_and_mesh_bitwise_exact':True,'independent_python_audit':row['independent_audit'],'seconds':seconds,'reduction_percent_vs_DLL4':100*(1-seconds['dll5']/seconds['dll4']),'final_field_sha256':row['field_sha256'],'final_mesh_sha256':row['mesh_sha256']}
 hsd=[]
 for name in sorted((ROOT/'results/hsd_regression').glob('*_auto.bin')):
  vba=name.with_name(name.name.replace('_auto.bin','_vba.bin'));old3=DLL3/'results/hsd_regression'/name.name
  assert name.read_bytes()==vba.read_bytes()==old3.read_bytes()
  hsd.append({'name':name.name,'actual_VBA_DLL3_DLL5_exact':True,'sha256':sha(name)})
 assert len(hsd)==6
 result={'q':8,'target':256,'FEM':records,'HSD_known_SOCP':hsd,'final_DLL5_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'passed':True,'scope':'fixed Fs, not final safety-factor search; all vector/mesh bytes compared, not just displayed values'}
 write(ROOT/'results/fixed_comparison.json',result);print('FINAL FIXED FULL-BINARY COMPARISON PASS',len(records),'HSD',len(hsd),flush=True)
if __name__=='__main__':run()
