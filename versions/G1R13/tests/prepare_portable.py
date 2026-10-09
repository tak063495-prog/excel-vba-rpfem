from pathlib import Path
import shutil
ROOT=Path(__file__).resolve().parents[1]
REPO=Path('C:/Users/link_/Downloads/excel-vba-rpfem_G1R11_20261008')
DIAG=Path('C:/Users/link_/Desktop/RPFEM_D03N_Trace_20261009')
for name in ['common.py','audit_precision.py']:
    src=REPO/'versions/G1R12/validation_runtime/python'/name;dst=ROOT/'validation_runtime/python'/name;dst.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(src,dst)
for name in ['independent_audit.py','zero_phi.py','model_support/__init__.py','model_support/schema.py','model_support/p2.py','model_support/upper_assembly.py','model_support/lower_assembly.py']:
    src=REPO/'versions/G1R12/validation_runtime/vendor'/name;dst=ROOT/'validation_runtime/vendor'/name;dst.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(src,dst)
dst=ROOT/'validation_runtime/data/cases.json';dst.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(REPO/'versions/G1R12/validation_runtime/data/cases.json',dst)
shutil.copy2(DIAG/'snapshot_reader.py',ROOT/'tests/snapshot_reader.py')
folder=ROOT/'evidence';folder.mkdir(exist_ok=True)
shutil.copy2(DIAG/'capture/candidate_s59_c1_i50.mesh.bin',folder/'D03N_original_rejected.mesh.bin')
shutil.copy2(DIAG/'capture/candidate_s59_c1_i50.metadata.tsv',folder/'D03N_original_rejected.metadata.tsv')
(ROOT/'requirements.txt').write_text('numpy==2.5.3\nscipy==1.18.1\nmpmath==1.4.1\nshapely==2.1.2\n',encoding='ascii')
print('PORTABLE RUNTIME READY')
