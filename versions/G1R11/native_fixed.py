from native_common import *
mode=sys.argv[1];assert mode in ['lower','upper']
fs,q=(.655,4) if mode=='lower' else (.77,8)
name='P06N_'+mode
work=R/'tests'/f'{name}_{time.time_ns()}.xlsm';shutil.copy2(source,work)
app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;book=None
try:
    book=app.Workbooks.Open(str(work),UpdateLinks=0);app.EnableEvents=False;inject(book)
    setvalue(book,'ADAPT_MESH_POLICY','LEGACY_THEN_REFINE');setvalue(book,'ADAPT_REFINE_CAP',1000);setvalue(book,'FS_BRACKET_POLICY','AUDITED_UPPER_LIMIT');setvalue(book,'QUADRATURE',q)
    compile_book(app,book);t=time.perf_counter();print('START',name,'native mesh order; Fs',fs,flush=True)
    out=app.Run(f"'{work.name}'!MeshSolve",str(R/'results/native1000_input.bin'),fs,mode,str(R/'results'/(name+'.bin')))
    rec={'test':name,'fs':fs,'mode':mode,'q':q,'result':out,'seconds':time.perf_counter()-t,'scope':'NATIVE_FIXED_FS_ON_NATIVE_REFINED_MESH_NOT_FULL_ADAPTATION'}
    (R/'results'/(name+'.json')).write_text(json.dumps(rec,ensure_ascii=False,indent=2),encoding='utf-8');print(rec,flush=True)
finally:
    if book is not None:book.Close(False)
    app.Quit()
