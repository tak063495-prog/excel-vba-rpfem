from native_repair import *
folder=ROOT/'controls';folder.mkdir(exist_ok=True)
template=ROOT/'cases/D07A/RPFEM_20261010_G1R13_D07A.xlsm'
app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
book=app.Workbooks.Open(str(template),UpdateLinks=0)
try:
    inject(book,'RPX_LoadStepMesh',(ROOT/'src/vba/RPX_LoadStepMesh.bas').read_text(encoding='utf-8'))
    compile_book(app,book);book.Save()
finally:book.Close(SaveChanges=False)
path=folder/'Repair_Controls.xlsm';assert not path.exists();shutil.copy2(template,path)
book=app.Workbooks.Open(str(path),UpdateLinks=0)
try:
    inject(book,'LoadStepControls',(ROOT/'tests/LoadStepControls.bas').read_text(encoding='utf-8'))
    compile_book(app,book)
    rows=[]
    for key in ['step','reversed','equal','zero','vertical','rotated','fixed','capacity','phase','cancel','subscript']:
        r=call(app,book,'LoadStepTest',key).split('\t');print(r,flush=True);rows.append(r)
    assert all(r[0]=='PASS' for r in rows[:7]),rows
    assert [r[2] for r in rows[7:]]==['7','5','18','9'],rows
    write(folder/'load_step_results.json',{'compile':'PASS','native_controls':rows,'pass':True})
    book.Save()
finally:book.Close(SaveChanges=False);app.Quit()
