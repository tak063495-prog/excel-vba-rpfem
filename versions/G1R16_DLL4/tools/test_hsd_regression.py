import gc,json,math,shutil,struct,time
import win32com.client
from prepare import ROOT,write,sha,REPO
from native_helpers import compile_book
def setvalue(book,key,value):
    ws=book.Worksheets('設定')
    for row in range(4,101):
        if ws.Cells(row,4).Value2==key:ws.Cells(row,2).Value2=value;return
    raise ValueError(key)
def inject(book,name):
    c=book.VBProject.VBComponents.Add(1);c.Name=name
    c.CodeModule.AddFromString('\r\n'.join(s for s in (ROOT/'tests'/f'{name}.bas').read_text(encoding='utf-8').splitlines() if not s.startswith('Attribute VB_')))
def run():
    folder=ROOT/'results/hsd_regression';folder.mkdir(parents=True,exist_ok=True)
    path=folder/'HSD_Regression.xlsm';shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm',path)
    shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
    app=win32com.client.DispatchEx('Excel.Application');book=None;result={'cases':[]}
    try:
        app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
        for name in ['HSDRegression','ContractTests','MeshProbe']:inject(book,name)
        compile_book(app,book)
        msg=str(app.Run("'"+book.Name+"'!TestLUContract"));print(msg,flush=True);result['LU_contract']=msg;assert msg.startswith('PASS')
        call=lambda name,*args:app.Run("'"+book.Name+"'!"+name,*args)
        for backend in ['LDL_CHECKED','PIVOT_ONLY']:
            setvalue(book,'HSD_LINEAR_BACKEND',backend)
            for case,(a,b,scalar) in enumerate([(0.6,0.8,False),(0.3,0.4,True),(6.,8.,True)]):
                record={'backend':backend,'case':case,'expected':math.hypot(a,b)+(.2 if scalar else 0)}
                for mode in ['vba','auto']:
                    call('RPX_NativeSetMode',mode);prefix=folder/f'{backend}_{case}_{mode}'
                    t=time.perf_counter();msg=str(call('HSDCase',a,b,scalar,str(prefix)));print(backend,case,mode,msg,flush=True)
                    record[mode]={'message':msg,'seconds':time.perf_counter()-t};assert msg.startswith('PASS'),msg
                    raw=prefix.with_suffix('.bin').read_bytes();record[mode]['objective']=struct.unpack_from('<d',raw)[0]
                    assert abs(record[mode]['objective']-record['expected'])<1e-7*(1+record['expected'])
                record['bitwise_exact']=(folder/f'{backend}_{case}_vba.bin').read_bytes()==(folder/f'{backend}_{case}_auto.bin').read_bytes()
                assert record['bitwise_exact'];result['cases'].append(record)
        book.Save()
    finally:
        write(folder/'summary.json',result)
        if book is not None:book.Close(False)
        app.Quit();book=None;app=None;gc.collect()
    print('HSD regression PASS',len(result['cases']),flush=True)
if __name__=='__main__':run()
