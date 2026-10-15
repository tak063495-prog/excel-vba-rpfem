"""Compile/reopen the single-DLL G1R19 workbook and verify every ZIP byte."""
from pathlib import Path
import gc,hashlib,json,shutil,time,zipfile
import psutil,win32com.client,win32process
ROOT=Path(__file__).resolve().parents[1];BASE=ROOT.parent/'G1R18';DELIVERY=ROOT/'delivery'
NAME='RPFEM_20261015_G1R19'
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def text(path):return '\r\n'.join(x for x in path.read_text(encoding='utf8').splitlines() if not x.startswith('Attribute VB_'))
def compile_book(app,book):
    book.Activate();book.VBProject.VBComponents('RPX_AccuracyNative').CodeModule.CodePane.Show()
    control=app.VBE.CommandBars.FindControl(1,578)
    if control.Enabled:control.Execute();time.sleep(.4)
    if control.Enabled:raise RuntimeError('G1R19 VBA compile failed')
    app.VBE.MainWindow.Visible=False
def create_book():
    DELIVERY.mkdir(parents=True,exist_ok=True);path=DELIVERY/(NAME+'.xlsm')
    if path.exists():raise FileExistsError('Refusing to overwrite an existing delivery workbook')
    shutil.copyfile(BASE/'delivery/RPFEM_20261014_G1R18.xlsm',path)
    assert len(list(DELIVERY.glob('*.dll')))==1
    assert sha(DELIVERY/'rpfem_native_x64.dll')==sha(ROOT/'native/unified/bin/rpfem_native_x64.dll')
    prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
    if prior:raise RuntimeError(f'Serial owned Excel required: {prior}')
    app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    record=dict(state='RUNNING',owned_excel_pid=pid,owned_excel_creation_time=psutil.Process(pid).create_time(),prior_pids=prior,modules={},dll_sha256=sha(DELIVERY/'rpfem_native_x64.dll'))
    save=lambda:(DELIVERY/'WORKBOOK_VALIDATION.json').write_text(json.dumps(record,indent=2,ensure_ascii=False)+'\n',encoding='utf8')
    save();book=component=control=None
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.MultiThreadedCalculation.Enabled=False
        book=app.Workbooks.Open(str(path),UpdateLinks=0)
        for name in [c.Name for c in book.VBProject.VBComponents if c.Type==1]:book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(name))
        for p in sorted((ROOT/'src/vba').glob('*.bas')):
            component=book.VBProject.VBComponents.Add(1);component.Name=p.stem;component.CodeModule.AddFromString(text(p))
            record['modules'][p.name]={'source_sha256':sha(p),'loaded_code_sha256':hashlib.sha256(component.CodeModule.Lines(1,component.CodeModule.CountOfLines).encode('utf8')).hexdigest()}
        assert len(record['modules'])==62;compile_book(app,book);record['compile']='PASS';save()
        call=lambda name,*values:app.Run("'"+book.Name+"'!"+name,*values)
        call('RPX_Setup')
        assert call('RPX_Setting','DIRECT_DLL',-1)==0 and call('RPX_Setting','GAP',-1)==.02
        component=book.VBProject.VBComponents.Add(1);component.Name='TwoPercentGuardProbe'
        component.CodeModule.AddFromString(text(ROOT/'tests/TwoPercentGuardProbe.bas'))
        record['two_percent_guard_unit']=call('TwoPercentGuardUnit')
        assert record['two_percent_guard_unit'].startswith('PASS;'),record['two_percent_guard_unit']
        book.VBProject.VBComponents.Remove(component);component=None
        inputs={n:book.Worksheets(n).UsedRange.Value2 for n in ['要素定義','材料データ','浸透入力']};policy=call('RPX_Setting','ANALYSIS_POLICY','')
        component=book.VBProject.VBComponents.Add(1);component.Name='DistributionPresetProbe'
        component.CodeModule.AddFromString('''
Public Function DistributionPresetProbeRun() As String
    Dim oldLower As RPX_MeshState, oldUpper As RPX_MeshState
    oldLower = RBestLower: oldUpper = RBestUpper
    On Error GoTo Failed
    ReDim RBestLower.field(0 To 1): ReDim RBestUpper.field(0 To 1)
    RBestLower.field(0) = 0.123456789012345: RBestLower.field(1) = -7.125
    RBestUpper.field(0) = 0.987654321098765: RBestUpper.field(1) = 9.25
    RPX_Speed2pctPreset
    If RBestLower.field(0) <> 0.123456789012345 Or RBestLower.field(1) <> -7.125 Then Err.Raise 5, , "PRESET_CHANGED_LOWER"
    If RBestUpper.field(0) <> 0.987654321098765 Or RBestUpper.field(1) <> 9.25 Then Err.Raise 5, , "PRESET_CHANGED_UPPER"
    DistributionPresetProbeRun = "PASS"
CleanUp:
    RBestLower = oldLower: RBestUpper = oldUpper
    Exit Function
Failed:
    DistributionPresetProbeRun = "FAIL " & Err.Description
    Resume CleanUp
End Function
''')
        assert call('DistributionPresetProbeRun')=='PASS'
        record['preset_preserves_saved_fields']=True
        book.VBProject.VBComponents.Remove(component);component=None;compile_book(app,book)
        expected={'GAP':.02,'DIRECT_DLL':1,'QUADRATURE':8,'FS_TOL':.0001,'ADAPT_CERT_QUADRATURE':8,'ADAPT_PRECISION_FROM_START':0,'ACCURACY_DLL_RECOVERY':1}
        record['speed_preset']={k:call('RPX_Setting',k,-1) for k in expected};assert record['speed_preset']==expected
        assert inputs=={n:book.Worksheets(n).UsedRange.Value2 for n in inputs} and policy==call('RPX_Setting','ANALYSIS_POLICY','')
        record['preset_preserves_input_and_policy']=True
        call('RPX_Setup');book.Worksheets('浸透結果').Cells.Clear();guide=book.Worksheets('実行案内');guide.Cells.Clear()
        for row,message in enumerate(['RPFEM G1R19','合格基準は上下界差2%です。','64bit Excelと同じフォルダー内のrpfem_native_x64.dll 1個を使用します。','設定シートの「2%・速度優先の設定」でDLL先行を選べます。','DLL先行は初期状態で無効です。原式監査の不合格時は従来解法へ戻ります。','同梱README_JA.mdの検証範囲と制限をご確認ください。'],1):guide.Cells(row,1).Value2=message
        guide.Columns('A').ColumnWidth=100;guide=None
        book.Worksheets('解析結果').Range('A1').Value2='G1R19 未解析（合格基準2%）'
        book.BuiltinDocumentProperties('Title').Value='RPFEM G1R19';book.Save();book.Close(False);book=None
        book=app.Workbooks.Open(str(path),UpdateLinks=0);compile_book(app,book)
        assert call('RPX_Setting','DIRECT_DLL',-1)==0 and call('RPX_Setting','GAP',-1)==.02
        assert call('RPX_Setting','ADAPT_PRECISION_FROM_START',-1)==0
        assert {c.Name for c in book.VBProject.VBComponents if c.Type==1}=={p.stem for p in (ROOT/'src/vba').glob('*.bas')}
        record.update(state='PASS',reopen_compile='PASS',initial_state='UNSOLVED;GAP=.02;DIRECT_DLL=0;precision_from_start=0;62_source_modules')
    except BaseException as error:record.update(state='FAIL',error=repr(error));raise
    finally:
        if book is not None:book.Close(False)
        app.Quit();book=component=control=app=None;gc.collect()
        for _ in range(50):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        record['owned_excel_exited']=not psutil.pid_exists(pid);record['workbook_sha256']=sha(path);save()
    return record
def package():
    files={}
    def add(name,path):
        if name in files:raise ValueError(name)
        files[name]=path.read_bytes()
    for name in [NAME+'.xlsm','rpfem_native_x64.dll','WORKBOOK_VALIDATION.json']:add(name,DELIVERY/name)
    for name in ['README_JA.md','SOURCE_MANIFEST.json','TEST_RESULTS.json','TEST_RESULTS.md']:
        if (ROOT/name).exists():add(name,ROOT/name)
    for directory in ['src/vba','import','tools','tests']:
        for p in sorted((ROOT/directory).rglob('*')):
            if p.is_file() and '__pycache__' not in p.parts:add(p.relative_to(ROOT).as_posix(),p)
    for p in sorted((ROOT/'native/unified').rglob('*')):
        relative=p.relative_to(ROOT/'native/unified')
        if not p.is_file() or any(x in relative.parts for x in ['target','obj','validation','.git']):continue
        if p.suffix.lower() in ['.dll','.lib','.obj','.pdb','.exe']:continue
        add(p.relative_to(ROOT).as_posix(),p)
    manifest={'release':'G1R19','acceptance_gap':.02,'runtime_requires_python':False,'runtime_dll_count':1,'p06_one_percent_goal_pursued':False,'byte_scope':'Exact generated release and ZIP bytes','files':{n:{'bytes':len(b),'sha256':hashlib.sha256(b).hexdigest()} for n,b in sorted(files.items())}}
    raw=(json.dumps(manifest,indent=2,ensure_ascii=False)+'\n').encode('utf8');(DELIVERY/'MANIFEST.json').write_bytes(raw);files['MANIFEST.json']=raw
    archive=DELIVERY/(NAME+'.zip')
    with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED,compresslevel=9) as out:
        for name,data in sorted(files.items()):
            info=zipfile.ZipInfo(NAME+'/'+name,(2026,10,15,0,0,0));info.compress_type=zipfile.ZIP_DEFLATED;out.writestr(info,data)
    with zipfile.ZipFile(archive) as check:
        for name,data in files.items():assert check.read(NAME+'/'+name)==data
        assert sum(n.lower().endswith('.dll') for n in check.namelist())==1
    report={'status':'PASS','archive_sha256':sha(archive),'bytes':archive.stat().st_size,'entries':len(files),'runtime_dll_count':1,'all_member_bytes_verified':True}
    (DELIVERY/'ZIP_VALIDATION.json').write_text(json.dumps(report,indent=2)+'\n',encoding='utf8');print(json.dumps(report),flush=True)
if __name__=='__main__':print(json.dumps(create_book(),ensure_ascii=False),flush=True);package()
