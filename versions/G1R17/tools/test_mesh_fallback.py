"""Actual native Excel G1R17 with missing/old DLL, no full analysis."""
import gc,json,shutil
import win32com.client
from support import ROOT,BASE,OUTPUT,write
from fixed_native import inject
from native_helpers import compile_book
def main():
 folder=ROOT/'results/mesh_fallback';folder.mkdir(exist_ok=False);rows=[]
 for mode,expected in [('missing','DLL_NOT_FOUND'),('old_dll','DLL_ABI_MISMATCH')]:
  sub=folder/mode;sub.mkdir();p=sub/'Fallback.xlsm';shutil.copyfile(ROOT/'delivery'/OUTPUT,p)
  if mode=='old_dll':shutil.copyfile(BASE/'delivery/rpfem_native_x64.dll',sub/'rpfem_native_x64.dll')
  app=win32com.client.DispatchEx('Excel.Application');book=None
  try:
   app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(p),UpdateLinks=0)
   for n in ['GeometryRegression','MeshProbe','HSDRegression']:inject(book,n)
   compile_book(app,book);call=lambda n,*a:app.Run("'"+book.Name+"'!"+n,*a)
   call('RPX_NativeSetMode','auto');info=str(call('RPX_NativeInfo'));assert info.startswith(expected),info
   prefix=sub/'P06N';msg=str(call('GeometryExport',str(prefix)));assert msg.startswith('PASS'),msg
   assert prefix.with_name(prefix.name+'_mesh.bin').read_bytes()==(ROOT/'results/geometry25/P06N_new_vba_mesh.bin').read_bytes()
   msg=str(call('HSDCase',.6,.8,False,str(sub/'hsd')));assert msg.startswith('PASS'),msg
   rows.append({'mode':mode,'native_reason':info,'mesh_bitwise_matches':True,'HSD_VBA_fallback':msg,'PASS':True});write(folder/'summary.json',{'rows':rows})
  finally:
   if book is not None:book.Close(False)
   app.Quit();book=None;app=None;gc.collect()
 write(folder/'summary.json',{'rows':rows,'PASS':True})
if __name__=='__main__':main()
