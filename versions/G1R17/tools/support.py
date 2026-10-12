from pathlib import Path
import hashlib,json
ROOT=Path(__file__).resolve().parents[1]
REPO=ROOT.parents[1]
BASE=REPO/'versions/G1R16_DLL5'
SOURCE=BASE/'delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm'
OUTPUT='RPFEM_20261012_G1R17_P06N_Coarse_Checked.xlsm'
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,v):
 p.parent.mkdir(parents=True,exist_ok=True)
 p.write_text(json.dumps(v,ensure_ascii=False,indent=2,default=lambda x:x.tolist() if hasattr(x,'tolist') else str(x))+'\n',encoding='utf8')
def setvalue(book,key,value):
 ws=book.Worksheets('設定')
 for r in range(4,101):
  if ws.Cells(r,4).Value2==key:ws.Cells(r,2).Value2=value;return
 for r in range(4,101):
  if ws.Cells(r,4).Value2 is None and ws.Cells(r,2).Value2 is None:ws.Cells(r,4).Value2=key;ws.Cells(r,2).Value2=value;return
 raise ValueError('No free setting slot '+key)
def configure(book,rec,target=200,q=8):
 app=book.Application;app.EnableEvents=False
 ws=book.Worksheets('要素定義');ws.Range('A7:T1006').ClearContents()
 for a,b,key in [('A','C','points'),('E','H','regions'),('J','T','faces')]:ws.Range(a+'7:'+b+str(6+len(rec[key]))).Value2=tuple(tuple(x)for x in rec[key])
 ws=book.Worksheets('材料データ');ws.Range('A5:Q1004').ClearContents();ws.Range('A5:E'+str(4+len(rec['materials']))).Value2=tuple(tuple(x)for x in rec['materials'])
 for key,value in {**rec['settings'],'GRAVITY':'固定','TARGET':target,'QUADRATURE':q,'ADAPT':0,'HSD_CAPTURE':0,'SEISMIC_KH':0,'HSD_SOLVE_SECONDS':300}.items():setvalue(book,key,value)
 book.Worksheets('浸透入力').Range('D4').Value2=0;app.EnableEvents=True

DLL4=REPO/"versions/G1R16_DLL4"
