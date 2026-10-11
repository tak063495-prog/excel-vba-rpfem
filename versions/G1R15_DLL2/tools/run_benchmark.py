from pathlib import Path
import argparse,json,time,shutil,hashlib,traceback
import win32com.client,win32process
ROOT=Path(__file__).resolve().parents[1]
parser=argparse.ArgumentParser(); parser.add_argument('--tag',default='dll_run1'); args=parser.parse_args()
folder=ROOT/'results'/args.tag; folder.mkdir(parents=True,exist_ok=False)
template=ROOT/'delivery/RPFEM_20261012_G1R15_DLL2_Levee_Checked.xlsm'
dst=folder/'RPFEM_G1R15_DLL_Levee.xlsm'
shutil.copyfile(template,dst); shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
app=win32com.client.DispatchEx('Excel.Application'); book=None; result={}
def write(): (folder/'status.json').write_text(json.dumps(result,ensure_ascii=False,indent=2,default=str),encoding='utf-8')
def call(name,*arguments): return app.Run("'"+book.Name+"'!"+name,*arguments)
try:
    app.Visible=False; app.DisplayAlerts=False
    book=app.Workbooks.Open(str(dst),UpdateLinks=0)
    result={'tag':args.tag,'excel_pid':win32process.GetWindowThreadProcessId(app.Hwnd)[1],
            'excel_version':app.Version,'security_unmodified':app.AutomationSecurity,
            'template_sha256':hashlib.sha256(template.read_bytes()).hexdigest(),
            'dll_sha256':hashlib.sha256((folder/'rpfem_native_x64.dll').read_bytes()).hexdigest(),
            'status':'preparing','settings':book.Worksheets('設定').Range('A1:F40').Value2,
            'seep_inputs':book.Worksheets('浸透入力').Range('A1:Q40').Value2,
            'geometry_inputs':book.Worksheets('要素定義').Range('A1:T40').Value2,
            'materials':book.Worksheets('材料データ').Range('A1:Q20').Value2}
    assert book.Worksheets('設定').Range('B5').Value2==600
    assert book.Worksheets('設定').Range('B7').Value2==1
    assert book.Worksheets('設定').Range('B15').Value2==.15
    assert book.Worksheets('浸透入力').Range('D4').Value2==1
    call('RPX_NativeSetMode','auto'); result['native_start']=str(call('RPX_NativeInfo'))
    assert result['native_start'].startswith('DLL_READY')
    call('RPX_SeepSetSilent',True)
    t=time.perf_counter(); result['seep_message']=str(call('RPX_SeepRunSilent')); result['seep_seconds']=time.perf_counter()-t
    t=time.perf_counter(); call('RPX_CreateMesh'); result['mesh_seconds']=time.perf_counter()-t
    result['status']='running'; write(); print('RUNNING',args.tag,'PID',result['excel_pid'],flush=True)
    t=time.perf_counter(); call('RPX_Run'); result['analysis_seconds']=time.perf_counter()-t
    result['native_end']=str(call('RPX_NativeInfo'))
    result['output']=book.Worksheets('解析結果').Range('A1:H40').Value2
    result['lower']=book.Worksheets('解析結果').Range('B4').Value2
    result['upper']=book.Worksheets('解析結果').Range('C4').Value2
    result['terminal_text']=book.Worksheets('解析結果').Range('A1').Value2
    result['status']='completed' if result['terminal_text']=='解析完了' else 'failed'
    # Persist outputs, then close the own instance; never kill unrelated Excel.
    book.Save(); result['output_sha256']=hashlib.sha256(dst.read_bytes()).hexdigest(); write()
    print(json.dumps({k:result[k] for k in ['status','analysis_seconds','lower','upper','native_end']},ensure_ascii=False),flush=True)
    assert result['status']=='completed',result['terminal_text']
finally:
    if not result:result={'status':'failed_before_open'}
    write()
    if book is not None:book.Close(False)
    app.Quit()
