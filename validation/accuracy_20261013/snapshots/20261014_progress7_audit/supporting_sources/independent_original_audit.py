"""Read-only independent audit of completed P06 witnesses and stored original guides."""
import hashlib
import json
from pathlib import Path
import sys

BASE = Path(r'D:\Work\excel-vba-rpfem\validation\accuracy_20261013')
sys.path.insert(0, str(BASE / 'python'))
import accuracy_driver as driver
import experiment_core as core
import report_accuracy as report

OUTPUT = BASE / 'results/progress7_independent_audit.json'
if OUTPUT.exists():
    raise ValueError('An independent output must be new')
implementation = {str(driver.REPO / p): hashlib.sha256((driver.REPO / p).read_bytes()).hexdigest()
                  for p in driver.SOURCE_HASHES}
entries = []
for case, mode, relative in [
    ('P06A', 'lower', 'a_lower22080_root_complete/P06A'),
    ('P06A', 'upper', 'a_upper12753_root_complete/P06A'),
    ('P06N', 'lower', 'n_lower11201_c0band_close/P06N'),
    ('P06N', 'upper', 'n_jump27917_fs6876_coordinate_close/P06N'),
]:
    folder = BASE / 'results' / relative
    _, record = driver.load_model(case)
    files = [folder / f'best_{mode}_{part}' for part in
             ('certificate.json', 'mesh.npz', 'field.npz')]
    files += sorted(folder.glob('*original_dual_guide_*.npz'))
    hashes = {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
    meta = json.loads(files[0].read_text(encoding='utf8'))
    mesh = driver.load_saved_mesh(files[1], record)
    with core.np.load(files[2], allow_pickle=False) as data:
        assert data['q'].shape == () and data['q'].item() == 8
        assert data['mode'].shape == () and str(data['mode']) == mode
        assert data['fs'].shape == ()
        fs, field = float(data['fs']), data['x'].copy()
    original = core.Prepared(mesh, fs, mode, 8, reduction='none')
    physical = core.evaluate(original, field)
    assert physical['physical_pass'] and physical['witness']
    assert report.root_is_complete(meta, mesh, fs, mode, case)
    original_guides = []
    for path in files[3:]:
        with core.np.load(path, allow_pickle=False) as data:
            guide = {k: data[k].item() for k in
                     ('case', 'mode', 'q', 'fs', 'mesh_id', 'physical_model_id', 'policy')}
            guide['dual'] = data['dual'].copy()
        guide_fs, evidence, _ = driver.audit_initial_dual_guide(mesh, mode, guide)
        root = meta['audit']['discrete_root_search']
        assert guide_fs == root['bracket'][1 if mode == 'lower' else 0]
        original_guides.append({'file': str(path), 'fs': guide_fs, 'audit': evidence})
    assert hashes == {str(p): hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
    entries.append(dict(case=case, mode=mode, fs=fs, mesh_id=mesh.mesh_id,
        physical_model_id=mesh.physical_model_id, policy=mesh.provenance['psi_policy'],
        original_primal=physical, modern_identity_and_signed_root_pass=True,
        stored_original_guides=original_guides,
        saved_full_negative_dual_pending=(case == 'P06N' and mode == 'lower' and not original_guides),
        input_hashes=hashes, input_hashes_unchanged=True))
    print('INDEPENDENT_PASS', case, mode, fs, 'guides', len(original_guides), flush=True)
assert implementation == {p: hashlib.sha256(Path(p).read_bytes()).hexdigest() for p in implementation}
core.write(OUTPUT, {'scope': 'Fresh original q8 primal and own-mesh root audit; saved original full duals audited when present',
                   'entries': entries, 'implementation_hashes': implementation,
                   'implementation_unchanged': True, 'modified_qp_duals_used': False,
                   'new_solver_used': False, 'strict_interval_arithmetic_proof': False})
