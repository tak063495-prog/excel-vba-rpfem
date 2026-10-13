import hashlib
import json
from pathlib import Path
import shutil
import sys

BASE = Path(r'D:\Work\excel-vba-rpfem\validation\accuracy_20261013')
sys.path.insert(0, str(BASE / 'python'))
import accuracy_driver as driver
import experiment_core as core
import original_dual

support = BASE / 'snapshots/20261014_progress7_audit/supporting_sources'
source = BASE / 'results/n_lower11201_completed_negative_full_support'
target = support / 'P06N/lower_full_negative_support'
if target.exists() and any(target.iterdir()):
    raise ValueError('A new support snapshot is required')
target.mkdir(exist_ok=True)
paths = [p for p in source.rglob('*') if p.is_file()]
before = {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
_, record = driver.load_model('P06N')
mesh = driver.load_saved_mesh(source / 'source_inputs/primal/best_lower_mesh.npz', record)
with core.np.load(source / 'original_helper_full_dual.npz', allow_pickle=False) as data:
    identity = {k: data[k].item() for k in ('case', 'mode', 'q', 'fs', 'mesh_id', 'physical_model_id', 'policy')}
    guide = dict(identity, dual=data['dual'].copy())
    source_hashes = json.loads(data['source_witness_sha256_json'].item())
assert identity == dict(case='P06N', mode='lower', q=8, fs=0.6783698418190449,
                       mesh_id=mesh.mesh_id, physical_model_id=mesh.physical_model_id,
                       policy=mesh.provenance['psi_policy'])
_, fresh, _ = driver.audit_initial_dual_guide(mesh, 'lower', guide)
assert fresh['accepted'] and fresh['search_sign'] == -1
for name, digest in source_hashes.items():
    assert hashlib.sha256((source / 'source_inputs/primal' / name).read_bytes()).hexdigest() == digest
for p in paths:
    destination = target / p.relative_to(source)
    destination.parent.mkdir(parents=True, exist_ok=True)
    shutil.copyfile(p, destination)
    assert hashlib.sha256(destination.read_bytes()).hexdigest() == before[str(p)]
assert before == {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in paths}
core.write(target / 'root_independent_full_dual_audit.json', {
    'identity': identity, 'fresh_original_full_dual_audit': fresh,
    'source_hashes': before, 'source_bytes_unchanged': True,
    'old_root_or_fields_modified': False, 'publishes_primal_bound': False})
shutil.copyfile(Path(__file__), target / 'root_independent_full_dual_audit.py')
manifest_path = support / 'manifest.json'
manifest = json.loads(manifest_path.read_text(encoding='utf8'))
manifest['n_lower_saved_full_negative_support'] = str(target.relative_to(support)).replace('\\', '/')
manifest['file_sha256'] = {str(p.relative_to(support)).replace('\\', '/'): hashlib.sha256(p.read_bytes()).hexdigest()
                          for p in sorted(support.rglob('*')) if p.is_file() and p != manifest_path}
core.write(manifest_path, manifest)
print('N_FULL_NEGATIVE_INDEPENDENT_PASS', fresh['dual_stationarity'], fresh['dual_cone_error'], fresh['dual_objective'], flush=True)
