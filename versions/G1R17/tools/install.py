"""Native Excel integration; original workbook and native DLL are preserved."""
import shutil,json,gc
import win32com.client
from support import ROOT,REPO,BASE,SOURCE,OUTPUT,write,sha,configure
from native_helpers import compile_book

def main():
 names=['RPX_Mesh','RPX_Native','RPX_NativeMesh']
 for name in names:
  src=ROOT/'src/vba'/(name+'.bas');s=src.read_text(encoding='utf8');data=s.replace('\n','\r\n').encode('cp932','strict')
  assert data.decode('cp932').replace('\r\n','\n')==s
  (ROOT/'import'/(name+'.bas')).write_bytes(data)
 p=ROOT/'delivery'/OUTPUT;assert not p.exists();shutil.copyfile(SOURCE,p)
 app=win32com.client.DispatchEx('Excel.Application');book=None
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(p),UpdateLinks=0)
  for name in names:
   if name!='RPX_NativeMesh':book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(name))
   book.VBProject.VBComponents.Import(str(ROOT/'import'/(name+'.bas')))
  compile_book(app,book)
  rec=json.loads((REPO/'validation/G1R12_native25_20261009/inputs/P06N.json').read_text(encoding='utf8'));configure(book,rec)
  book.Worksheets('要素定義').Range('A3').Value2='P06N・Davis等価関連・G1R17・DLL5改訂粗メッシュ修復。TARGET200、q8、適応OFF。c=0自由頂点の鈍角ファンのみ局所分割。'
  book.Save();assert book.VBProject.VBComponents.Count==71
  write(ROOT/'results/install_g1r17.json',{'compile':'PASS','module_count':71,'changed_modules':names,'native_DLL_unchanged':sha(ROOT/'delivery/rpfem_native_x64.dll')==sha(BASE/'delivery/rpfem_native_x64.dll'),'source_sha256':sha(SOURCE),'candidate_sha256':sha(p),'security_unmodified':app.AutomationSecurity})
 finally:
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
 print('INSTALL COMPILE PASS',flush=True)
if __name__=='__main__':main()
