"""Fixed Fs only, P06A/P06N VBA vs G1R17 DLL. No Fs search/adaptation."""
import gc,json,shutil,time,struct
import win32com.client
from support import ROOT,OUTPUT,write,sha,configure
from fixed_native import inject
from native_helpers import compile_book
def main():
 folder=ROOT/'results/fixed_precision';folder.mkdir(exist_ok=False);path=folder/'Precision.xlsm';rows=[]
 shutil.copyfile(ROOT/'delivery'/OUTPUT,path);shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
 app=win32com.client.DispatchEx('Excel.Application');book=None
 try:
  app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
  for n in ['MeshProbe','HSDRegression']:inject(book,n)
  compile_book(app,book);call=lambda n,*a:app.Run("'"+book.Name+"'!"+n,*a)
  for case,fs in [('P06N',.5),('P06A',.7)]:
   configure(book,json.loads((ROOT/'data/inputs'/(case+'.json')).read_text(encoding='utf8')))
   row={'case':case,'Fs':fs,'side':'lower'};rows.append(row)
   for mode in ['vba','auto']:
    call('RPX_NativeSetMode',mode);prefix=folder/(case+'_'+mode);print('START FIXED',case,fs,mode,flush=True)
    t=time.perf_counter();msg=str(call('FixedFem','lower',fs,str(prefix)));assert msg.startswith('PASS'),msg
    row[mode]={'result':msg,'seconds':time.perf_counter()-t};write(folder/'summary.json',{'rows':rows})
    print('END FIXED',case,mode,msg,flush=True)
   row['mesh_bitwise_exact']=(folder/(case+'_vba_mesh.bin')).read_bytes()==(folder/(case+'_auto_mesh.bin')).read_bytes()
   row['field_bitwise_exact']=(folder/(case+'_vba_field.bin')).read_bytes()==(folder/(case+'_auto_field.bin')).read_bytes()
   assert row['mesh_bitwise_exact']and row['field_bitwise_exact'],row
   if case=='P06N':
    row['old_DLL5_repaired_mesh_field_bitwise_exact']=(ROOT/'results/fixed_native/lower_0.5_field.bin').read_bytes()==(folder/(case+'_auto_field.bin')).read_bytes();assert row['old_DLL5_repaired_mesh_field_bitwise_exact']
   write(folder/'summary.json',{'rows':rows})
  write(folder/'summary.json',{'rows':rows,'PASS':True,'DLL_sha256':sha(folder/'rpfem_native_x64.dll')})
 finally:
  if book is not None:book.Close(False)
  app.Quit();book=None;app=None;gc.collect()
if __name__=='__main__':main()
