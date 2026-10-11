from pathlib import Path
import json,shutil,sys,time,traceback,hashlib,argparse
import win32com.client
ROOT=Path(__file__).resolve().parents[1]
from native_helpers import compile_book

def call(app,book,name,*args): return app.Run("'"+book.Name+"'!"+name,*args)
def inject(book,name,text):
    component=book.VBProject.VBComponents.Add(1); component.Name=name
    component.CodeModule.AddFromString('\r\n'.join(s for s in text.splitlines() if not s.startswith('Attribute VB_')))

source=ROOT.parent/'G1R15/examples/RPFEM_20261011_G1R15_Levee_Example.xlsm'
dst=ROOT/'delivery/RPFEM_20261012_G1R15_DLL2_Levee_Checked.xlsm'
parser=argparse.ArgumentParser(); parser.add_argument('--update',action='store_true'); args=parser.parse_args()
if args.update:
    backup=ROOT/'results'/('template_before_update_'+str(time.time_ns())+'.xlsm')
    assert dst.exists() and not backup.exists(), 'Preserve the prior working template once'
    shutil.copyfile(dst,backup)
else:
    assert not dst.exists() or hashlib.sha256(dst.read_bytes()).digest()==hashlib.sha256(source.read_bytes()).digest(), 'Do not overwrite an existing modified delivery'
    shutil.copyfile(source,dst)
app=win32com.client.DispatchEx('Excel.Application'); book=None
results={}
try:
    app.Visible=False; app.DisplayAlerts=False
    book=app.Workbooks.Open(str(dst),UpdateLinks=0)
    for name in ['RPX_LDL','RPX_SchurLDL','RPX_SeepUI','RPX_Order','RPX_ColumnOrder']:
        book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(name))
        book.VBProject.VBComponents.Import(str(ROOT/'import'/f'{name}.bas'))
    if args.update:book.VBProject.VBComponents.Remove(book.VBProject.VBComponents('RPX_Native'))
    book.VBProject.VBComponents.Import(str(ROOT/'import/RPX_Native.bas'))
    inject(book,'NativeTests',(ROOT/'tests/NativeTests.bas').read_text(encoding='utf-8'))
    compile_book(app,book)
    for name in ['RPX_NativeInfo','RPX_TestHSDPrecision','TestNativeKernels','TestNativePivot','TestNativeCancellation','TestNativeShortBuffer','TestNativeOverflow','TestNativeOrdering','TestOrderingCancellation','TestOrderingBadIndex']:
        results[name]=str(call(app,book,name)); print(name,results[name],flush=True)
        assert results[name].startswith(('PASS','DLL_READY'))
    # Same settings as the existing 622-second run; seepage is run separately.
    settings=book.Worksheets('設定')
    settings.Range('B5').Value2=600
    settings.Range('B7').Value2=1
    settings.Range('B15').Value2=0.15
    settings.Range('B16').Value2='右(+x)'
    book.Worksheets('浸透入力').Range('D4').Value2=1
    # Test code is kept outside the delivered project.
    book.VBProject.VBComponents.Remove(book.VBProject.VBComponents('NativeTests'))
    compile_book(app,book)
    book.Save()
finally:
    (ROOT/'results/native_kernel_tests.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
    if book is not None: book.Close(False)
    app.Quit()
missing=ROOT/'tests/missing_dll'; missing.mkdir(exist_ok=True)
missing_book=missing/'Missing_DLL.xlsm'; shutil.copyfile(dst,missing_book)
app=win32com.client.DispatchEx('Excel.Application'); book=None
try:
    app.Visible=False; app.DisplayAlerts=False
    book=app.Workbooks.Open(str(missing_book),UpdateLinks=0)
    inject(book,'NativeTests',(ROOT/'tests/NativeTests.bas').read_text(encoding='utf-8'))
    result=str(call(app,book,'TestMissingDLL'))
    print(result,flush=True); assert result.startswith('PASS')
    results['missing_dll']=result
finally:
    (ROOT/'results/native_kernel_tests.json').write_text(json.dumps(results,ensure_ascii=False,indent=2),encoding='utf-8')
    if book is not None: book.Close(False)
    app.Quit()
