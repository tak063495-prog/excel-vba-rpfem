import shutil,json,time,gc,win32com.client
from prepare import ROOT,SOURCE,sha,write
from native_helpers import compile_book
path=ROOT/'delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm'
if path.exists():
    backup=ROOT/'results'/('working_before_update_'+str(time.time_ns())+'.xlsm');shutil.copyfile(path,backup)
shutil.copyfile(SOURCE,path)
app=win32com.client.DispatchEx('Excel.Application');book=None
try:
    app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
    modules=json.loads((ROOT/'results/import_modules.json').read_text(encoding='utf-8'))['modules']
    names={c.Name for c in book.VBProject.VBComponents}
    for name in modules:
        if name in names:book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(name))
        book.VBProject.VBComponents.Import(str(ROOT/'import'/f'{name}.bas'))
    compile_book(app,book)
    info=str(app.Run("'"+book.Name+"'!RPX_NativeInfo"));print(info,flush=True);assert info.startswith('DLL_READY')
    book.Save();write(ROOT/'results/install.json',{'compile':'PASS','security_unmodified':app.AutomationSecurity,'modules':modules,'native_info':info,'source_sha256':sha(SOURCE),'delivery_sha256':sha(path)})
except Exception:
    if book is not None:
        cm=book.VBProject.VBComponents('RPX_NativeBulk').CodeModule
        (ROOT/'results/compile_bulk_actual.bas').write_text(cm.Lines(1,cm.CountOfLines),encoding='utf-8')
        print('ACTUAL BULK HOP',cm.ProcStartLine('RPX_NativeHOperator',0),flush=True)
    raise
finally:
    if book is not None:book.Close(False)
    app.Quit();book=None;app=None;gc.collect()
