"""Close an own-mesh root using fresh original primal/dual vector audits."""
import os,sys,argparse,json,hashlib
from pathlib import Path
for n in ['OPENBLAS_NUM_THREADS','MKL_NUM_THREADS','OMP_NUM_THREADS']:os.environ[n]='1'
sys.path.insert(0,r'D:\Work\excel-vba-rpfem\validation\accuracy_20261013\python')
import accuracy_driver as driver
import experiment_core as core
p=argparse.ArgumentParser();p.add_argument('--case',required=True);p.add_argument('--physical',type=Path,required=True)
p.add_argument('--dual',type=Path,required=True);p.add_argument('--output',type=Path,required=True)
a=p.parse_args()
if a.output.exists():raise ValueError('New output required')
_,rec=driver.load_model(a.case)
mesh=driver.load_saved_mesh(a.physical/'best_upper_mesh.npz',rec)
metadata=json.loads((a.physical/'best_upper_certificate.json').read_text(encoding='utf8'))
with core.np.load(a.physical/'best_upper_field.npz',allow_pickle=False) as data:
    field,fs=data['x'].copy(),float(data['fs']);assert str(data['mode'])=='upper' and int(data['q'])==8
expected={'case':a.case,'mode':'upper','q':8,'fs':fs,'mesh_id':mesh.mesh_id,'physical_model_id':mesh.physical_model_id}
assert all(metadata.get(k)==v for k,v in expected.items())
pr=core.Prepared(mesh,fs,'upper',8,reduction='none');audit=core.evaluate(pr,field)
assert audit['physical_pass'] and audit['witness']
dual_meta=json.loads((a.dual/'audit.json').read_text(encoding='utf8'))
dual_mesh=driver.load_saved_mesh(a.dual/'original_mesh.npz',rec)
assert dual_meta['case']==a.case and dual_meta['q']==8 and dual_mesh.mesh_id==mesh.mesh_id
assert dual_meta['mesh_id']==mesh.mesh_id and dual_meta['physical_model_id']==mesh.physical_model_id
with core.np.load(a.dual/'original_candidate.npz',allow_pickle=False) as data:dual=data['z'].copy()
dual_fs=float(dual_meta['fs'])
assert 0<dual_fs<fs and fs-dual_fs<=1e-4*max(1.,(dual_fs+fs)/2)
guide={'case':a.case,'mode':'upper','q':8,'fs':dual_fs,'dual':dual,'mesh_id':mesh.mesh_id,
    'physical_model_id':mesh.physical_model_id,'policy':mesh.provenance['psi_policy']}
driver.audit_initial_dual_guide(mesh,'upper',guide)
audit.update(expected)
warm=driver.Certificate('upper',fs,mesh,field,audit,'fresh original primal-vector audit',False)
a.output.mkdir(parents=True);(a.output/'scratch_source.py').write_bytes(Path(__file__).read_bytes())
best,root=driver.search(mesh,'upper',(dual_fs,fs),a.output,a.case+'_original_pair_close',1e-4,120,
    initial_witness=warm,initial_dual_guides=[guide])
assert root['root_complete'] and best.root_complete and best.fs==fs
metadata=best.publish(a.output)
core.write(a.output/'import_audit.json',{'case':a.case,'certificates':{'upper':metadata}})
core.write(a.output/'summary.json',{'case':a.case,'upper':metadata,'upper_only':True,
    'one_percent_target_not_evaluated':True,'root':root,'implementation_hashes':driver.SOURCE_HASHES})
sources={str(path):hashlib.sha256(path.read_bytes()).hexdigest() for path in
    [a.physical/'best_upper_mesh.npz',a.physical/'best_upper_field.npz',a.physical/'best_upper_certificate.json',
     a.dual/'original_candidate.npz',a.dual/'original_mesh.npz',a.dual/'audit.json']}
core.write(a.output/'source_provenance.json',{'original_primal_source':str(a.physical),
    'original_dual_source':str(a.dual),'source_files_sha256':sources,'field_or_root_flags_transferred':False})
print('CLOSED_ORIGINAL_UPPER',a.case,fs,root['bracket'],root['bracket_width'],flush=True)
