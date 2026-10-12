from pathlib import Path
import hashlib,json,sys,math
ROOT=Path(__file__).resolve().parents[1]
REPO=ROOT.parents[1]
BASE=REPO/'versions/G1R16_DLL5'
SOURCE=BASE/'delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,v):
 p.parent.mkdir(parents=True,exist_ok=True)
 def clean(x):
  if isinstance(x,dict):return {k:clean(v)for k,v in x.items()}
  if isinstance(x,(list,tuple)):return [clean(v)for v in x]
  if isinstance(x,float)and not math.isfinite(x):return None
  if hasattr(x,'tolist'):return clean(x.tolist())
  return x
 p.write_text(json.dumps(clean(v),ensure_ascii=False,indent=2,allow_nan=False,default=str)+'\n',encoding='utf8')
sys.path.insert(0,str(ROOT/'validation_runtime/python'))
