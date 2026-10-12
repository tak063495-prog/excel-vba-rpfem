from pathlib import Path
import hashlib,json,shutil,subprocess,sys
ROOT=Path(__file__).resolve().parents[1]
REPO=ROOT.parents[1]
DLL3=REPO/'versions/G1R16_DLL_3'
DLL4=REPO/'versions/G1R16_DLL4'
SOURCE=DLL4/'delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm'
PYTHON=Path(sys.executable)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,value):
 p.parent.mkdir(parents=True,exist_ok=True)
 p.write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
def main():
 assert not (ROOT/'baseline_sources').exists()
 names=subprocess.check_output(['git','ls-files','versions/G1R16_DLL4'],cwd=REPO,text=True).splitlines()
 for name in names:
  p=REPO/name;rel=p.relative_to(DLL4)
  if rel.parts[0] not in ['src','native','tests','import','tools'] and p.name not in ['.gitignore','.gitattributes']:continue
  if rel==Path('tools/prepare.py'):continue
  dest=ROOT/rel;dest.parent.mkdir(parents=True,exist_ok=True)
  if p.suffix=='.py' or rel.parts[0]=='tests':
   s=p.read_text(encoding='utf-8').replace('G1R16_DLL4','G1R16_DLL5').replace('DLL4','DLL5')
   dest.write_text(s,encoding='utf-8',newline='\n')
  else:shutil.copyfile(p,dest)
 shutil.copytree(DLL4/'src/vba',ROOT/'baseline_sources')
 for folder in ['delivery','results']:(ROOT/folder).mkdir(exist_ok=True)
 protected={str(p.relative_to(REPO)).replace('\\','/'):sha(p) for folder in ['versions/G1R16','versions/G1R15_DLL2','versions/G1R16_DLL_3','versions/G1R16_DLL4'] for p in (REPO/folder).rglob('*') if p.is_file()}
 write(ROOT/'results/baseline.json',{'source':str(SOURCE.relative_to(REPO)),'source_sha256':sha(SOURCE),'source_commit':subprocess.check_output(['git','rev-parse','HEAD'],cwd=REPO,text=True).strip(),'module_count':69,'protected':protected})
 write(ROOT/'WORK_STATUS.json',{'state':'IMPLEMENTING','source':'G1R16 DLL4 commit6f9c224','authority':'User requested complete DLL5 optimization without postponing improvements to DLL6. Existing P06 full/extreme exploration stop preferences remain in force.','planned':['semantic exact input comparison','full Schur numerical path','HSD DD assembly and Krylov repair','time-sliced bulk jobs and call fusion','compiled model preparation/structure reuse','precision, failure injection and lifecycle regression','DLL4/VBA speed and precision comparison','native Excel saved workbook, ZIP and authorized GitHub delivery'],'not_run':['DLL5 precision','DLL5 robustness','DLL5 performance','DLL5 publication']})
 print('DLL5 prepared;',len(protected),'original files protected')
if __name__=='__main__':main()
