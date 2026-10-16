"""Fresh owned Excel, test-only traps; no solver/field/root certificates."""
from pathlib import Path
import gc,hashlib,importlib.util,json,shutil,threading,time,uuid
import psutil,win32com.client,win32con,win32gui,win32process

ROOT=Path(__file__).resolve().parents[1];REPO=ROOT.parents[1]
BASE=REPO/'versions/G1R19';STATE=REPO/'versions/G1R19_StateReset_20261016';WATER=REPO/'versions/G1R19_WaterCommon_20261016'
STATE_SHA='a26beb0538dfccd33021c4ea603a702957a2452b44ed08dc23bfbe75b088a147'
WATER_SHA='e25c0748e7a0c1b8127a2baa540d29448fe6d788ed04691cd601e931946e2f5c'
spec=importlib.util.spec_from_file_location('state_probes',STATE/'tests/make_probe_sources.py')
probes=importlib.util.module_from_spec(spec);spec.loader.exec_module(probes)
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def code(p):return '\r\n'.join(s for s in p.read_text(encoding='utf8').splitlines() if not s.startswith('Attribute VB_'))

PREP='''
' TEST-ONLY setup, explicit Call prevents a parameterless call becoming a label.
Public Function SRHarnessPrepare() As String
    On Error GoTo Failed
    Call RPX_InputEnd
    XCancel = False: AFActive = False: RFsActive = False: RProgramKind = ""
    Call RPX_InputBegin
    Call RPX_ReadInputs
    DTarget = 16
    Call RPX_GenerateMesh
    Call RPX_TotalReferenceLoads
    RFsActive = True
    SRHarnessPrepare = "PASS fresh_dry_D02_mesh;q8;solver0"
    Exit Function
Failed:
    SRHarnessPrepare = "FAIL " & Err.Number & " " & Err.Description
End Function
Public Sub SRHarnessEnd()
    RFsActive = False: XCancel = False
    Call RPX_InputEnd
End Sub
'''

def run(negative):
    assert sha(STATE/'SOURCE_MANIFEST.json')==STATE_SHA
    assert sha(WATER/'SOURCE_MANIFEST.json')==WATER_SHA
    prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
    if prior:raise RuntimeError('Another Excel exists: '+repr(prior))
    out=ROOT/'results'/('negative_' if negative else 'positive_')/str(time.time_ns());out.mkdir(parents=True)
    src=out/'test_sources';probes.generate(src,negative)
    # Build the same ten-module product composition as the future candidate,
    # retaining test-only StateReset instrumentation and its separate manifest.
    for p in (WATER/'src/vba').glob('*.bas'):
        shutil.copyfile(p,src/'src/vba'/p.name)
        shutil.copyfile(WATER/'import'/p.name,src/'import'/p.name)
    setup=src/'src/vba/StateResetMicroProbe.bas'
    text=setup.read_text(encoding='utf8').rstrip()+'\n'+PREP
    for sub,enc in [('src/vba','utf8'),('import','cp932')]:
        (src/sub/setup.name).write_bytes(text.replace('\r\n','\n').replace('\n','\r\n').encode(enc))
    files={p.stem:p for p in (src/'src/vba').glob('*.bas')};assert len(files)==63
    bookpath=out/'StateResetMicro.xlsm';shutil.copyfile(BASE/'delivery/RPFEM_20261015_G1R19.xlsm',bookpath)
    shutil.copyfile(BASE/'delivery/rpfem_native_x64.dll',out/'rpfem_native_x64.dll')
    rec=dict(status='PREPARING',test_only=True,numerical_evidence_allowed=False,negative_control=negative,
        token=uuid.uuid4().hex,source_sha256={n:sha(p) for n,p in files.items()},
        state_manifest_sha256=STATE_SHA,water_manifest_sha256=WATER_SHA,harness_sha256=sha(__file__),
        dll_sha256=sha(out/'rpfem_native_x64.dll'),prior_excel_pids=prior,Fs_solver_calls=0,dialogs=[],results={})
    save=lambda:(out/'summary.json').write_text(json.dumps(rec,indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8')
    save();shutil.copyfile(__file__,out/'harness_snapshot.py')
    app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    rec.update(owned_excel_pid=pid,owned_excel_creation_time=psutil.Process(pid).create_time());save()
    stop=threading.Event()
    def watch():
        while not stop.wait(.2):
            def visit(w,_):
                if win32process.GetWindowThreadProcessId(w)[1]!=pid or win32gui.GetClassName(w)!='#32770':return
                kids=[];win32gui.EnumChildWindows(w,lambda c,_:kids.append((c,win32gui.GetWindowText(c))),None)
                words=[t for _,t in kids]
                if words not in rec['dialogs']:rec['dialogs'].append(words)
                for c,t in kids:
                    if t in ['OK','&OK'] or ('Visual Basic' in win32gui.GetWindowText(w) and ('終了' in t or '&End' in t)):win32gui.PostMessage(c,win32con.BM_CLICK,0,0)
            win32gui.EnumWindows(visit,None)
    watcher=threading.Thread(target=watch,daemon=True);watcher.start();book=c=control=None
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.MultiThreadedCalculation.Enabled=False
        book=app.Workbooks.Open(str(bookpath),UpdateLinks=0)
        for name in [c.Name for c in book.VBProject.VBComponents if c.Type==1]:book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(name))
        for name,p in files.items():
            c=book.VBProject.VBComponents.Add(1);c.Name=name;c.CodeModule.AddFromString(code(p))
        c.CodeModule.CodePane.Show();control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled:control.Execute();time.sleep(.3)
        if control.Enabled:
            pane=app.VBE.ActiveCodePane;line=int(pane.GetSelection()[0])
            raise RuntimeError('VBA compile failed '+pane.CodeModule.Name+': '+pane.CodeModule.Lines(max(1,line-2),5))
        app.VBE.MainWindow.Visible=False;rec['compile']='PASS'
        rec['loaded_code_sha256']={c.Name:hashlib.sha256(c.CodeModule.Lines(1,c.CodeModule.CountOfLines).encode('utf8')).hexdigest() for c in book.VBProject.VBComponents if c.Type==1};save()
        call=lambda n,*v:app.Run("'"+book.Name+"'!"+n,*v)
        call('RPX_Setup')
        path=REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs/D02A.json'
        data=json.loads(path.read_text(encoding='utf8'));rec['input_sha256']=sha(path);shutil.copyfile(path,out/'input_record.json')
        for sheet,ranges in [('要素定義',[('A','C','points',7),('E','H','regions',7),('J','T','faces',7)]),('材料データ',[('A','E','materials',5)])]:
            ws=book.Worksheets(sheet);ws.Range('A7:T1006' if sheet=='要素定義' else 'A5:Q1004').ClearContents()
            for first,last,key,row in ranges:
                values=tuple(tuple(x) for x in data[key]);ws.Range(f'{first}{row}:{last}{row+len(values)-1}').Value2=values
            ws=None
        def setting(k,v):
            ws=book.Worksheets('設定')
            values=ws.Range('B4:D1000').Value2
            for row,record in enumerate(values,4):
                if record[2]==k:ws.Cells(row,2).Value2=v;return
            for row,record in enumerate(values,4):
                if record[2] in [None,''] and record[0] in [None,'']:
                    ws.Cells(row,4).Value2=k;ws.Cells(row,2).Value2=v;return
            raise KeyError(k)
        for k,v in data['settings'].items():setting(k,v)
        call('RPX_Speed2pctPreset');setting('TARGET',16)
        def test(key,name,*args):
            value=call(name,*args);rec['results'][key]=value;save()
            if negative:
                assert str(value).startswith('FAIL'),(key,value)
                reasons={'conic':'GUIDE_ENTRY_DID_NOT_CLEAR_SENTINEL','direct_lower':'DIRECT_ENTRY_DID_NOT_RESET_PRIVATE_SENTINELS','direct_upper':'DIRECT_ENTRY_DID_NOT_RESET_PRIVATE_SENTINELS','krylov':'TEST_ONLY_KRYLOV_ENTRY_GUARD_WAS_SKIPPED','seep':'defaults=False'}
                assert reasons[key] in str(value),(key,value)
            else:assert str(value).startswith('PASS'),(key,value)
        test('conic','SRProbeConicEntry')
        for side in ['lower','upper']:
            rec['results']['prepare_'+side]=call('SRHarnessPrepare');save();assert rec['results']['prepare_'+side].startswith('PASS')
            test('direct_'+side,'SRProbeDirectEntry',side,1.3);call('SRHarnessEnd')
        assert call('SRHarnessPrepare').startswith('PASS')
        test('krylov','SRProbeKrylovEntry');call('SRHarnessEnd')
        test('seep','SRProbeSeepEntry')
        rec['native_memory_maps']=[m.path for m in psutil.Process(pid).memory_maps() if Path(m.path).name.lower().startswith('rpfem_')]
        assert rec['native_memory_maps']==[str(out/'rpfem_native_x64.dll')]
        rec.update(status='PASS',meaning='Old entry skips reproduced; no numerical certificates' if negative else 'Explicit entry calls execute, all traps before solvers')
        book.Save();save()
    except BaseException as e:rec.update(status='FAIL',error=repr(e));save();raise
    finally:
        try:
            if book is not None:book.Close(False)
        finally:app.Quit()
        stop.set();watcher.join(2);book=c=control=app=None;gc.collect()
        for _ in range(50):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        rec['owned_excel_exited']=not psutil.pid_exists(pid);save()
        rec['files']={str(p.relative_to(out)):{'bytes':p.stat().st_size,'sha256':sha(p)} for p in out.rglob('*') if p.is_file() and p.name!='summary.json'};save()
    assert sha(STATE/'SOURCE_MANIFEST.json')==STATE_SHA and sha(WATER/'SOURCE_MANIFEST.json')==WATER_SHA
    print(json.dumps({'status':rec['status'],'negative_control':negative,'output':str(out),'owned_excel_exited':rec['owned_excel_exited']}),flush=True)
    return out

if __name__=='__main__':
    run(False)
    run(True)
