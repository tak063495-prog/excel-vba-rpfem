import json
from prepare import ROOT,DLL3,write,sha

def run():
 result={'fixed_fem':[],'HSD':[],'comparison':'actual recorded DLL3 outputs against actual paired VBA/DLL4 outputs; bit-for-bit including binary metadata'}
 for name in sorted((ROOT/'results/fixed_fem_final').glob('*_auto_field.bin')):
  old=DLL3/'results/fixed_fem_final'/name.name;vba=name.with_name(name.name.replace('_auto_','_vba_'))
  assert name.read_bytes()==old.read_bytes()==vba.read_bytes()
  mesh=name.with_name(name.name.replace('_field.bin','_mesh.bin'));oldmesh=old.with_name(old.name.replace('_field.bin','_mesh.bin'));vbamesh=vba.with_name(vba.name.replace('_field.bin','_mesh.bin'))
  assert mesh.read_bytes()==oldmesh.read_bytes()==vbamesh.read_bytes()
  result['fixed_fem'].append({'file':name.name,'vba_dll3_dll4_bitwise_exact':True,'DLL3_field_sha256':sha(old),'DLL4_field_sha256':sha(name),'same_mesh':True})
 for name in sorted((ROOT/'results/hsd_regression').glob('*_auto.bin')):
  old=DLL3/'results/hsd_regression'/name.name;vba=name.with_name(name.name.replace('_auto.bin','_vba.bin'))
  assert name.read_bytes()==old.read_bytes()==vba.read_bytes()
  result['HSD'].append({'file':name.name,'vba_dll3_dll4_bitwise_exact':True,'DLL3_sha256':sha(old),'DLL4_sha256':sha(name)})
 assert len(result['fixed_fem'])==len(result['HSD'])==6
 result['note']='q2 fixed FEM was executed before final clock/budget/guard-only changes; final q8 three-backend FEM is recorded separately. HSD canonical cases use the final DLL and workbook.'
 write(ROOT/'results/three_way_precision.json',result);print('THREE WAY q2/FEM and canonical HSD PASS 12',flush=True)
if __name__=='__main__':run()
