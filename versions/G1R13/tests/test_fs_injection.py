from native_repair import *
folder=ROOT/'fs_controls';folder.mkdir(exist_ok=True)
path=folder/'Fs_Controls_Retry.xlsm';assert not path.exists();shutil.copy2(ROOT/'cases/D07A/RPFEM_20261010_G1R13_D07A.xlsm',path)
app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;book=None
try:
    book=app.Workbooks.Open(str(path),UpdateLinks=0)
    s=(ROOT/'src/vba/RPX_Fs.bas').read_text(encoding='utf-8')+'\n'+(ROOT/'tests/fs_injection_snippet.bas').read_text(encoding='utf-8')
    inject(book,'RPX_Fs',s);compile_book(app,book)
    ws=book.Worksheets('設定')
    for row in range(4,101):
        if ws.Cells(row,4).Value2=='TEST_INJECTION':ws.Cells(row,2).Value2=1;break
    cases=[(5,'HSD_DD_PRODUCT_UNDERFLOW'),(7,'HSD_DD_PRODUCT_UNDERFLOW'),(9,'HSD_DD_PRODUCT_UNDERFLOW'),(18,'ANALYSIS_CANCELLED'),(5,'UPPER_PHYSICAL_AUDIT_FAILED injected'),(5,'UNKNOWN_ERROR'),(5,'HSD_DD_PRODUCT_UNDERFLOW extra')]
    rows=[]
    for number,message in cases:
        r=call(app,book,'FsRepairInjectedTrial',number,message).split('\t');print(r,flush=True);rows.append(r)
    assert rows[0]==['UNKNOWN','NUMERICAL_UNKNOWN','','1','123'],rows
    assert [r[1] for r in rows[1:]]==['7','9','18','5','5','5'],rows
    assert all(r[-1]=='1' for r in rows[1:]) and rows[0][3]=='1',rows
    ws.Cells(row,2).Value2=0
    write(folder/'result.json',{'compile':'PASS','controlled_failure_injection':'actual production TryFsValue / fs_value hook; no numerical solver execution','cases':rows,'pass':True})
    book.Save()
finally:
    if book is not None:book.Close(SaveChanges=False)
    app.Quit();gc.collect()
