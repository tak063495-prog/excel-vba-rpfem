from pathlib import Path
import json,subprocess,sys
import win32com.client
ROOT=Path(__file__).resolve().parents[1]
app=win32com.client.DispatchEx('Excel.Application'); book=None
try:
    app.Visible=False; app.DisplayAlerts=False
    book=app.Workbooks.Open(str(ROOT/'delivery/RPFEM_20261012_G1R15_DLL2_Levee_Checked.xlsm'),UpdateLinks=0,ReadOnly=True)
    data={'native_info':app.Run("'"+book.Name+"'!RPX_NativeInfo"),
          'lower':book.Worksheets('解析結果').Range('B4').Value2,
          'upper':book.Worksheets('解析結果').Range('C4').Value2,
          'status':book.Worksheets('解析結果').Range('A1').Value2,
          'macro_security_unmodified':app.AutomationSecurity}
    assert data['native_info'].startswith('DLL_READY') and data['status']=='解析完了'
    (ROOT/'results/saved_delivery_open.json').write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(data,ensure_ascii=False))
finally:
    if book is not None:book.Close(False)
    app.Quit()
baseline=Path('C:/Users/link_/Downloads/RPFEM_Codex_P2S_0945_Task')
result=subprocess.run([sys.executable,'-X','utf8','tools/verify_baseline.py'],cwd=baseline,capture_output=True,text=True,encoding='utf-8',check=True)
data=json.loads(result.stdout); assert data['baseline_intact']
(ROOT/'results/historical_baseline_preserved.json').write_text(json.dumps(data,indent=2),encoding='utf-8')
