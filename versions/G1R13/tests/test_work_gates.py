from native_control_helpers import *
folder=ROOT/'work_controls';folder.mkdir(exist_ok=True)
path=folder/'Work_Gates.xlsm';assert not path.exists();shutil.copy2(ROOT/'single_point/D03N/RPFEM_20261010_G1R13_D03N.xlsm',path)
app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;book=None
try:
    book=app.Workbooks.Open(str(path),UpdateLinks=0)
    for p in (ROOT/'src/vba').glob('*.bas'):inject(book,p.stem,p.read_text(encoding='utf-8'))
    inject(book,'WorkGateControls',(ROOT/'tests/WorkGateControls.bas').read_text(encoding='utf-8'));compile_book(app,book)
    seq=tuple(r['value'] for r in json.loads((ROOT/'evidence/factor_reconstruction.json').read_text(encoding='utf-8'))['selected'])
    r=call(app,book,'D03ReplayPrepare',str(ROOT/'evidence/D03N_original_rejected.mesh.bin'),str(folder/'raw_problem.bin'),seq,str(ROOT/'evidence/model_aux.csv'),str(ROOT/'evidence/expected_model.txt'))
    assert r.startswith('READY '),r
    call(app,book,'WorkGateRead',str(ROOT/'evidence/D03N_original_rejected.mesh.bin'))
    cases=['original_repair','zero','negative','large_work','numerical_gate','wrong_mode','certificate','positive_c','fixed_load','positive_cost','rigid','raw_phase','broken_equation','field_dimensions','slack_dimensions']
    rows=[]
    for case in cases:
        r=call(app,book,'WorkGateTest',case).split('\t');rows.append(r);print(r,flush=True)
    assert all(r[0]=='PASS' for r in rows[:13]),rows
    assert all(r[2]=='9' for r in rows[13:]),rows
    write(folder/'result.json',{'compile':'PASS','pass':True,'scope':'native production normalization function on actual original rejected mesh/field; model/goal limits, transactional rejects, zero dual, dimensions','cases':rows})
    book.Save()
finally:
    if book is not None:book.Close(SaveChanges=False)
    app.Quit();gc.collect()
