from pathlib import Path
import json,shutil
import win32com.client
ROOT=Path(__file__).resolve().parents[1]
folder=ROOT/'tests/missing_order'; folder.mkdir(exist_ok=True)
dst=folder/'Missing_DLL_Order.xlsm'
shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R15_DLL2_Levee_Checked.xlsm',dst)
app=win32com.client.DispatchEx('Excel.Application'); book=None
try:
    app.Visible=False; app.DisplayAlerts=False
    book=app.Workbooks.Open(str(dst),UpdateLinks=0)
    component=book.VBProject.VBComponents.Add(1); component.Name='NativeTests'
    component.CodeModule.AddFromString('\r\n'.join(s for s in (ROOT/'tests/NativeTests.bas').read_text(encoding='utf-8').splitlines() if not s.startswith('Attribute VB_')))
    result=str(app.Run("'"+book.Name+"'!TestMissingDLL")); assert result.startswith('PASS'),result
    p=ROOT/'results/native_kernel_tests.json'; data=json.loads(p.read_text(encoding='utf-8'))
    data['missing_dll']=result; p.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8')
    print(result)
finally:
    if book is not None:book.Close(False)
    app.Quit()
