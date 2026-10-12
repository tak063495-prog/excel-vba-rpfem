from pathlib import Path
import hashlib,json,shutil,sys,subprocess
ROOT=Path(__file__).resolve().parents[1]
REPO=ROOT.parents[1]
DLL3=REPO/'versions/G1R16_DLL_3'
SOURCE=DLL3/'delivery/RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm'
PYTHON=Path(sys.executable)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,value):
 p.parent.mkdir(parents=True,exist_ok=True)
 p.write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
def prepare():
 assert not (ROOT/'baseline_sources').exists()
 tracked=subprocess.check_output(['git','ls-files','versions/G1R16_DLL_3'],cwd=REPO,text=True).splitlines()
 for name in tracked:
  p=REPO/name; rel=p.relative_to(DLL3)
  if rel.parts[0] in ['src','native','tests','import','tools'] or p.name in ['.gitignore','.gitattributes']:
   if rel==Path('tools/prepare.py'):continue
   dest=ROOT/rel;dest.parent.mkdir(parents=True,exist_ok=True);shutil.copyfile(p,dest)
   if rel.parts[0]=='tools' and p.suffix=='.py':
    s=dest.read_text(encoding='utf-8').replace('G1R16_DLL_3','G1R16_DLL4').replace('G1R16_DLL_3','G1R16_DLL4').replace('DLL_3','DLL4')
    dest.write_text(s,encoding='utf-8',newline='\n')
 shutil.copytree(DLL3/'src/vba',ROOT/'baseline_sources')
 for folder in ['delivery','results','docs']:(ROOT/folder).mkdir(exist_ok=True)
 protected={str(p.relative_to(REPO)).replace('\\','/'):sha(p) for folder in ['versions/G1R16','versions/G1R15_DLL2','versions/G1R16_DLL_3'] for p in (REPO/folder).rglob('*') if p.is_file()}
 write(ROOT/'results/baseline.json',{'source':str(SOURCE.relative_to(REPO)),'source_sha256':sha(SOURCE),'source_commit':subprocess.check_output(['git','rev-parse','HEAD'],cwd=REPO,text=True).strip(),'dll3_workbook_sha256':sha(DLL3/'delivery/RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm'),'module_count':68,'protected':protected})
 write(ROOT/'WORK_STATUS.json',{'state':'IMPLEMENTING','scope':'DLL4: persistent factors, fused direction/refinement/vector operations for normal IPM and HSD; compare VBA/DLL3 and native robustness, complete searches and timing before publication','not_run':['precision','robustness','speed','publication']})
 print('DLL4 prepared; original protected:',len(protected))
if __name__=='__main__':prepare()
