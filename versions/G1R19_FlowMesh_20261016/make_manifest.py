"""Encode/pin only this new pure geometry module; no existing source mutation."""
from pathlib import Path
import hashlib,json
ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[1]

def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def main():
    source=ROOT/'src/vba/RPX_FlowMesh.bas'
    text=source.read_text(encoding='utf8').replace('\r\n','\n').replace('\r','\n').replace('\n','\r\n')
    source.write_bytes(text.encode('utf8'))
    imported=ROOT/'import/RPX_FlowMesh.bas';imported.parent.mkdir(exist_ok=True);imported.write_bytes(text.encode('cp932'))
    tracked=[ROOT/'README_JA.md',ROOT/'tests/FlowMeshArrayProbe.bas',ROOT/'tests/reference_fixtures.py',
             ROOT/'tests/test_contract.py',ROOT/'tests/compare_excel.py',Path(__file__)]
    fixture_root=ROOT/'tests/fixtures'
    if (fixture_root/'FIXTURES.json').exists():tracked+=sorted(p for p in fixture_root.iterdir() if p.is_file())
    reference=REPO/'validation/G1R19_python26_20261016/python/mesh_geometry.py'
    manifest=dict(schema='rpfem.flowmesh.pure_array.v1',status='STATIC_AND_PYTHON_REFERENCE_ONLY; VBA_COMPILE_AND_RUNTIME_PENDING',
                  production_artifacts_modified=False,solver_calls=0,excel_runs=0,
                  module=dict(source='src/vba/RPX_FlowMesh.bas',source_sha256=sha(source),imported='import/RPX_FlowMesh.bas',import_sha256=sha(imported),decoded_text_equal=source.read_text(encoding='utf8')==imported.read_text(encoding='cp932')),
                  reference=dict(path=reference.relative_to(REPO).as_posix(),sha256=sha(reference)),
                  scope='Pure geometry candidate only; no global model/field/Fs/root access; new original q8 witness and own roots required',
                  semantics=dict(default_passes=5,source_metric='Fixed source centroid P2 J.T@J by original element ID',
                                 spectral_floor='eigenvalues >= max(lambda_max,1e-20)/9; normalize by max(lambda_max,1e-20)',
                                 metric_quality='after > 1.05*before+1e-10',geometric_quality='after >= .6*before',
                                 cavity_twice_area_absolute_tolerance=1e-10,minimum_positive_twice_area=1e-12,
                                 material_area='abs(delta) <= 1e-10+1e-12*abs(source_area)',
                                 extra_guards='Explicit RConstraints plus intrinsic exterior/material edge protection',
                                 ordering='Descending gain; stable original edge discovery ordinal for exact ties',
                                 limits='0..5 passes; exact finite 12*ne physical velocity prefix only'),
                  implementation_difference='Symmetric 2x2 spectral floor uses an analytic polynomial rather than an eigenvector routine; rounding may affect near-tied geometric choices. Extra explicit-convexity/existing-diagonal guards reject invalid cavities.',
                  files={p.relative_to(ROOT).as_posix():sha(p) for p in tracked})
    (ROOT/'SOURCE_MANIFEST.json').write_text(json.dumps(manifest,indent=2,ensure_ascii=False)+'\n',encoding='utf8')
    print(json.dumps(dict(status=manifest['status'],source_sha256=sha(source),manifest_sha256=sha(ROOT/'SOURCE_MANIFEST.json'))))

if __name__=='__main__':main()
