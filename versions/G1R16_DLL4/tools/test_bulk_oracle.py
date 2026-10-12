import gc,shutil,time,json,sys
import win32com.client
from prepare import ROOT,write,sha
from native_helpers import compile_book
def run():
    folder=ROOT/'results/bulk_oracle';folder.mkdir(parents=True,exist_ok=True)
    path=folder/'Bulk_Oracle.xlsm';shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm',path)
    shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
    app=win32com.client.DispatchEx('Excel.Application');book=None;results={}
    try:
        app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
        book.VBProject.VBComponents('RPX_Solve').CodeModule.AddFromString((ROOT/'tests/BulkInstrumentation.txt').read_text(encoding='utf-8'))
        for name in ['BulkOracle','NativeTests']:
            c=book.VBProject.VBComponents.Add(1);c.Name=name
            c.CodeModule.AddFromString('\r\n'.join(s for s in (ROOT/'tests'/f'{name}.bas').read_text(encoding='utf-8').splitlines() if not s.startswith('Attribute VB_')))
        compile_book(app,book)
        for mode in ['vba','auto']:
            t=time.perf_counter();msg=app.Run("'"+book.Name+"'!DumpBulkOracle",str(folder/(mode+'.bin')),mode)
            results[mode]={'message':msg,'seconds':time.perf_counter()-t};print(mode,msg,flush=True)
            assert not msg.startswith('FAIL'),msg
        for name in ['RPX_TestHSDPrecision','TestNativeKernels','TestNativePivot','TestNativeCancellation','TestNativeShortBuffer','TestNativeOverflow','TestNativeOrdering','TestOrderingCancellation','TestOrderingBadIndex']:
            if '--bulk-only' in sys.argv and name=='TestNativeOrdering':continue
            msg=str(app.Run("'"+book.Name+"'!"+name));print(name,msg,flush=True);results[name]=msg;assert msg.startswith('PASS'),msg
        book.Save()
    finally:
        if book is not None:book.Close(False)
        app.Quit();book=None;app=None;gc.collect()
    a=(folder/'vba.bin').read_bytes();b=(folder/'auto.bin').read_bytes();differences=[i for i,(x,y) in enumerate(zip(a,b)) if x!=y]
    results.update({'bytes_vba':len(a),'bytes_native':len(b),'different_bytes':len(differences),'first_differences':differences[:25],'exact':a==b,'dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll')})
    write(folder/'summary.json',results);print(json.dumps(results,ensure_ascii=False),flush=True)
    assert a==b
if __name__=='__main__':run()
