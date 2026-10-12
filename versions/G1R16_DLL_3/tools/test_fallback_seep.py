import gc,json,shutil
import win32com.client
from prepare import ROOT,REPO,write
from native_helpers import compile_book
from test_hsd_regression import inject,setvalue
def run():
    result={}
    for kind in ['native','missing','old_dll']:
        folder=ROOT/'results'/('fallback_'+kind);folder.mkdir(parents=True,exist_ok=True)
        path=folder/'Fallback.xlsm';shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm',path)
        if kind!='missing':shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll' if kind=='native' else REPO/'versions/G1R15_DLL2/delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
        app=win32com.client.DispatchEx('Excel.Application');book=None;rows={}
        try:
            app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
            for name in ['HSDRegression','MeshProbe','NativeTests']:inject(book,name)
            compile_book(app,book);call=lambda name,*args:app.Run("'"+book.Name+"'!"+name,*args)
            call('RPX_NativeSetMode','auto');rows['native_info']=str(call('RPX_NativeInfo'));print(kind,rows['native_info'],flush=True)
            assert rows['native_info'].startswith('DLL_READY') if kind=='native' else not rows['native_info'].startswith('DLL_READY')
            setvalue(book,'HSD_LINEAR_BACKEND','PIVOT_ONLY');rows['HSD']=str(call('HSDCase',.6,.8,True,str(folder/'fallback_hsd')));assert rows['HSD'].startswith('PASS')
            if kind=='missing':rows['missing_kernels']=str(call('TestMissingDLL'));assert rows['missing_kernels'].startswith('PASS')
            if kind=='native':
                rows['dam']=str(call('RPX_SeepTestDam',40,20));assert 'iters=54 converged=True Q=2.4181' in rows['dam']
                for relax in [1e-8,.5,1.]:
                    text=str(call('RPX_SeepTestRelax',relax));print('relax',relax,text,flush=True);rows[str(relax)]=text
                    assert ('iters=300;converged=False' if relax==1e-8 else 'iters=21;converged=True' if relax==.5 else 'iters=2;converged=True') in text
                rows['water_table']=str(call('RPX_SeepWaterTableSelfTest'));assert rows['water_table'].count('合格')==2
            result[kind]=rows;print(kind,'PASS',flush=True)
        finally:
            if book is not None:book.Close(False)
            app.Quit();book=None;app=None;gc.collect()
            write(ROOT/'results/fallback_seep.json',result)
    reference=(ROOT/'results/fallback_native/fallback_hsd.bin').read_bytes()
    assert reference==(ROOT/'results/fallback_missing/fallback_hsd.bin').read_bytes()==(ROOT/'results/fallback_old_dll/fallback_hsd.bin').read_bytes()
    result['same_hsd_result_all_backends']=True;write(ROOT/'results/fallback_seep.json',result)
if __name__=='__main__':run()
