from pathlib import Path
import shutil,json,hashlib,subprocess,sys
R=Path(__file__).resolve().parent
OLD=Path('C:/Users/link_/Desktop/RPFEM_G1R10_Mesh_20261007')
for name in ['src','original','tests','results','import_cp932','rollback_cp932']:(R/name).mkdir(exist_ok=True)
for p in (OLD/'src').glob('*.bas'):
    shutil.copy2(p,R/'src'/p.name);shutil.copy2(p,R/'original'/p.name)
for p in ['inspect_book.py','python_mesh_experiments.py','requirements-validation.txt','tests/native_helpers.py']:
    shutil.copy2(OLD/p,R/p)
if not (R/'validation_runtime').exists():shutil.copytree(OLD/'validation_runtime',R/'validation_runtime',ignore=shutil.ignore_patterns('__pycache__'))
source=OLD/'RPFEM_20261007_G1R10_P06N_Robust_Mesh.xlsm'
sha=hashlib.sha256(source.read_bytes()).hexdigest();assert sha=='98002b027f1d8b38a758423998b409d32798e00fd0aeb073e3bcae3c3cc01f8d'
(R/'source_identity.json').write_text(json.dumps({'path':str(source),'sha256':sha},indent=2),encoding='utf-8')
tool=Path('C:/Users/link_/Downloads/RPFEM_Codex_P2S_0945_Task/tools/verify_baseline.py')
p=subprocess.run([sys.executable,str(tool)],cwd=tool.parent.parent,capture_output=True,text=True,encoding='utf-8',check=True)
(R/'results/baseline_before.json').write_text(p.stdout,encoding='utf-8')
print('Protected source copied; old baseline verification:',p.stdout,flush=True)
