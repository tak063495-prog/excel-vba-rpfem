import gc,json,time,shutil,traceback,argparse
import win32com.client,win32process
from prepare import ROOT,write,sha
parser=argparse.ArgumentParser();parser.add_argument('--tag',default='levee_native');args=parser.parse_args()
assert args.tag.isidentifier()
folder=ROOT/'results'/args.tag;folder.mkdir(parents=True,exist_ok=False)
source=ROOT/'delivery/RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm'
path=folder/'DLL_3_Levee_Result.xlsm';shutil.copyfile(source,path)
shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
app=win32com.client.DispatchEx('Excel.Application');book=None;result={}
def call(name,*args):return app.Run("'"+book.Name+"'!"+name,*args)
try:
    app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
    settings=book.Worksheets('設定');seep=book.Worksheets('浸透入力')
    result={'status':'preparing','excel_pid':win32process.GetWindowThreadProcessId(app.Hwnd)[1],'excel_version':app.Version,'security_unmodified':app.AutomationSecurity,'source_sha256':sha(source),'dll_sha256':sha(folder/'rpfem_native_x64.dll'),'settings':settings.Range('A1:F40').Value2,'seep_inputs':seep.Range('A1:Q40').Value2,'geometry':book.Worksheets('要素定義').Range('A1:T40').Value2,'materials':book.Worksheets('材料データ').Range('A1:Q20').Value2}
    assert settings.Range('B5').Value2==600 and settings.Range('B7').Value2==1 and settings.Range('B15').Value2==.15 and seep.Range('D4').Value2==1
    call('RPX_NativeSetMode','auto');result['native_start']=str(call('RPX_NativeInfo'));assert result['native_start'].startswith('DLL_READY')
    call('RPX_SeepSetSilent',True)
    t=time.perf_counter();result['seep_message']=str(call('RPX_SeepRunSilent'));result['seep_seconds']=time.perf_counter()-t
    result['seep_outputs']=book.Worksheets('浸透結果').Range('J1:X6').Value2
    t=time.perf_counter();call('RPX_CreateMesh');result['mesh_seconds']=time.perf_counter()-t
    result['status']='running';write(folder/'status.json',result);print('RUNNING PID',result['excel_pid'],flush=True)
    t=time.perf_counter();call('RPX_Run');result['analysis_seconds']=time.perf_counter()-t
    result['output']=book.Worksheets('解析結果').Range('A1:H40').Value2
    result['lower']=book.Worksheets('解析結果').Range('B4').Value2;result['upper']=book.Worksheets('解析結果').Range('C4').Value2
    result['terminal_text']=book.Worksheets('解析結果').Range('A1').Value2
    result['native_end']=str(call('RPX_NativeInfo'));result['status']='completed' if result['terminal_text']=='解析完了' else 'failed'
    book.Save();result['output_sha256']=sha(path);write(folder/'status.json',result)
    print(json.dumps({k:result[k] for k in ['status','analysis_seconds','lower','upper','native_end']},ensure_ascii=False),flush=True)
    assert result['status']=='completed'
except Exception:
    result['exception']=traceback.format_exc();result['status']='failed';raise
finally:
    write(folder/'status.json',result)
    if book is not None:book.Close(False)
    app.Quit();book=None;app=None;gc.collect()
