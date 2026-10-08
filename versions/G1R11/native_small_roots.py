from native_common import *
rows=[]
for policy in ['LEGACY','AUDITED_UPPER_LIMIT']:
    work=R/'tests'/f'SmallRoot_{policy}_{time.time_ns()}.xlsm';shutil.copy2(source,work)
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;book=None
    try:
        book=app.Workbooks.Open(str(work),UpdateLinks=0);app.EnableEvents=False;inject(book)
        setvalue(book,'ANALYSIS_POLICY','ASSOCIATED');setvalue(book,'FS_BRACKET_POLICY',policy)
        book.Worksheets('材料データ').Range('E5:E7').Value2=((35.,),(25.,),(35.,))
        cm=book.VBProject.VBComponents.Add(1);cm.Name='SmallRootProbe';cm.CodeModule.AddFromString('\r\n'.join((R/'tests/SmallRootProbe.bas').read_text().splitlines()[1:]))
        compile_book(app,book);t=time.perf_counter()
        prefix=R/'results'/f'C5_final_{policy}'
        out=app.Run(f"'{work.name}'!SmallRoots",str(prefix))
        rec={'policy':policy,'result':out,'prefix':str(prefix),'seconds':time.perf_counter()-t,'scope':'ACTUAL_NATIVE_FS_ROOTS_C5_PHI_PSI45_SMALL_SLOPE_NOT_P06_FULL_RUN'}
        rows.append(rec);(R/'results/native_small_roots.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8');print(rec,flush=True)
        assert out.startswith('PASS;'),out
    finally:
        if book is not None:book.Close(False)
        app.Quit()
a=float(rows[0]['result'].split('lower=')[1].split(';')[0]);b=float(rows[1]['result'].split('lower=')[1].split(';')[0])
assert abs(a-b)<2e-4*max(1,abs(a),abs(b))
print('PASS actual old/new small Fs roots within root-width tolerance; audits required separately',flush=True)
