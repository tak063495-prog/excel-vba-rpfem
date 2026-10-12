from pathlib import Path
import hashlib,json,shutil,sys
from oletools.olevba import VBA_Parser
ROOT=Path(__file__).resolve().parents[1]
REPO=ROOT.parents[1]
SOURCE=REPO/'versions/G1R16/delivery/RPFEM_20261012_G1R16_DLL2_Levee_Checked.xlsm'
PYTHON=Path(sys.executable)
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,value):
    p.parent.mkdir(parents=True,exist_ok=True)
    p.write_text(json.dumps(value,ensure_ascii=False,indent=2),encoding='utf-8')
def prepare():
    assert not (ROOT/'baseline_sources').exists()
    for folder in ['baseline_sources','src/vba','native','delivery','tests','results','import','docs']:(ROOT/folder).mkdir(parents=True,exist_ok=True)
    parser=VBA_Parser(str(SOURCE))
    try:
        for _,_,name,code in parser.extract_macros():
            text=code.replace('\r\n','\n')
            for directory in ['baseline_sources','src/vba']:(ROOT/directory/name).write_text(text,encoding='utf-8',newline='\n')
    finally:parser.close()
    assert len(list((ROOT/'src/vba').glob('*')))==66
    old=REPO/'versions/G1R15_DLL2'
    for p in (old/'native').iterdir():
        if p.suffix in ['.c','.cpp','.asm','.cmd']:shutil.copyfile(p,ROOT/'native'/p.name)
    for name in ['NativeTests.bas','order_failure_injection.cpp','build_failure_test.cmd']:
        shutil.copyfile(old/'tests'/name,ROOT/'tests'/name)
    shutil.copyfile(old/'tools/native_helpers.py',ROOT/'tools/native_helpers.py')
    protected={str(p.relative_to(REPO)).replace('\\','/'):sha(p) for folder in ['versions/G1R16','versions/G1R15_DLL2'] for p in (REPO/folder).rglob('*') if p.is_file()}
    write(ROOT/'results/baseline.json',{'source':str(SOURCE.relative_to(REPO)),'source_sha256':sha(SOURCE),'source_commit':'6ab56ff4c8662089c8768c26f9d3b580b58f07ad','module_count':66,'protected':protected})
    print('G1R16 source extracted: 66 modules; original releases protected',len(protected))
if __name__=='__main__':prepare()
