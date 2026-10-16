"""Owned Excel pure-array geometry comparison; no model installation or solver."""
from pathlib import Path
import gc,hashlib,importlib.util,json,shutil,sys,threading,time,uuid
import psutil,win32com.client,win32con,win32gui,win32process

ROOT=Path(__file__).resolve().parents[1];REPO=ROOT.parents[1]
FLOW=REPO/'versions/G1R19_FlowMesh_20261016'
CANDIDATE=REPO/'versions/G1R19_Combined_20261016/results/candidate_book_v1'
FLOW_SHA='32fe459e26590730d59a368cd2cdc2528b770420136c4457a9a998ee12ffa75b'
BOOK_SHA='0ab98297993536b8b40372e36d440f27a6b1cfee8ccad648a2f5e149b08722ef'
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def code(p):return '\r\n'.join(s for s in p.read_text(encoding='utf8').splitlines() if not s.startswith('Attribute VB_'))

SCOPE='''Attribute VB_Name = "FlowScopeProbe"
Option Explicit
Public Function FlowScopePrepare() As String
    On Error GoTo Failed
    Call RPX_InputBegin
    Call RPX_ReadInputs
    DTarget = 16
    Call RPX_GenerateMesh
    Call RPX_TotalReferenceLoads
    FlowScopePrepare = "PASS fresh_global_geometry;solver0"
    Exit Function
Failed:
    FlowScopePrepare = "FAIL " & Err.Number & " " & Err.Description
End Function
Private Function FlowScopeArray(ByRef a() As Double) As String
    Dim i As Long, high As Long
    On Error GoTo Missing
    high = UBound(a)
    On Error GoTo 0
    For i = LBound(a) To high: FlowScopeArray = FlowScopeArray & RPX_ExactDouble(a(i)): Next i
    Exit Function
Missing:
    If Err.Number <> 9 Then Err.Raise Err.Number, Err.Source, Err.Description
    FlowScopeArray = "UNALLOCATED"
End Function
Public Function FlowScopeKey() As String
    Dim a As Double, b As Double, lr As Boolean, ur As Boolean
    lr = RPX_ActualRootBracket("lower", a, b)
    ur = RPX_ActualRootBracket("upper", a, b)
    FlowScopeKey = RPX_ExactModel() & "|" & RPX_InputEpoch & "|" & RPX_LoadPhase & "|" & RPX_AnalysisPolicy()
    FlowScopeKey = FlowScopeKey & "|" & RSubdivisions & "|" & RPX_ExactDouble(RStrengthFactor) & "|" & RFsCalls & "|" & RFsActive
    FlowScopeKey = FlowScopeKey & "|" & RPX_ExactDouble(RLowerFs) & "|" & RPX_ExactDouble(RUpperFs) & "|" & lr & "|" & ur
    FlowScopeKey = FlowScopeKey & "|" & RPX_ExactDouble(RBestLower.value) & "|" & RPX_ExactDouble(RBestUpper.value)
    FlowScopeKey = FlowScopeKey & "|" & XN & "|" & XM & "|" & XB & "|" & XS & "|" & RProgramKind & "|" & RPX_ExactDouble(RProgramFs)
    FlowScopeKey = FlowScopeKey & "|" & XCancel & "|" & XStatus & "|" & RPX_ExactDouble(XObjective) & "|" & RPX_ExactDouble(RValue)
    FlowScopeKey = FlowScopeKey & "|" & RPX_ResultCurrent & "|" & RPX_MeshCurrent & "|" & RPhysicalGateAccepted & "|" & RPX_AuditRejected
    FlowScopeKey = FlowScopeKey & "|" & FlowScopeArray(xx) & "|" & FlowScopeArray(xy) & "|" & FlowScopeArray(XSlack) & "|" & FlowScopeArray(XDual)
End Function
Public Sub FlowScopeEnd()
    Call RPX_InputEnd
End Sub
'''

def main():
    assert sha(FLOW/'SOURCE_MANIFEST.json')==FLOW_SHA
    original=CANDIDATE/'delivery/RPFEM_20261016_G1R19_Combined_candidate.xlsm'
    assert sha(original)==BOOK_SHA
    prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
    if prior:raise RuntimeError('Another Excel exists: '+repr(prior))
    out=ROOT/'results'/('owned_geometry_'+str(time.time_ns()));out.mkdir(parents=True)
    bookpath=out/'FlowMeshGeometry.xlsm';shutil.copyfile(original,bookpath)
    shutil.copyfile(CANDIDATE/'delivery/rpfem_native_x64.dll',out/'rpfem_native_x64.dll')
    files={'RPX_FlowMesh':FLOW/'src/vba/RPX_FlowMesh.bas','FlowMeshArrayProbe':FLOW/'tests/FlowMeshArrayProbe.bas'}
    fixture_root=FLOW/'tests/fixtures';fixture_manifest=json.loads((fixture_root/'FIXTURES.json').read_text(encoding='utf8'))
    record=dict(status='PREPARING',token=uuid.uuid4().hex,solver_calls=0,certificate=False,model_installation=False,
        input_candidate_book_sha256=BOOK_SHA,flow_manifest_sha256=FLOW_SHA,harness_sha256=sha(__file__),
        source_sha256={n:sha(p) for n,p in files.items()},fixture_manifest_sha256=sha(fixture_root/'FIXTURES.json'),
        base_source_manifest_sha256=sha(CANDIDATE/'SOURCE_MANIFEST.json'),prior_excel_pids=prior,dialogs=[],fixtures={})
    save=lambda:(out/'summary.json').write_text(json.dumps(record,indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8')
    save();shutil.copyfile(__file__,out/'harness_snapshot.py');(out/'FlowScopeProbe.bas').write_text(SCOPE,encoding='utf8')
    for n,p in files.items():shutil.copyfile(p,out/p.name)
    shutil.copyfile(FLOW/'SOURCE_MANIFEST.json',out/'FLOW_SOURCE_MANIFEST.json')
    app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    record.update(owned_excel_pid=pid,owned_excel_creation_time=psutil.Process(pid).create_time());save()
    stop=threading.Event()
    def watch():
        while not stop.wait(.2):
            def visit(w,_):
                if win32process.GetWindowThreadProcessId(w)[1]!=pid or win32gui.GetClassName(w)!='#32770':return
                kids=[];win32gui.EnumChildWindows(w,lambda c,_:kids.append((c,win32gui.GetWindowText(c))),None)
                if [t for _,t in kids] not in record['dialogs']:record['dialogs'].append([t for _,t in kids])
                for c,t in kids:
                    if t in ['OK','&OK'] or ('Visual Basic' in win32gui.GetWindowText(w) and ('終了' in t or '&End' in t)):win32gui.PostMessage(c,win32con.BM_CLICK,0,0)
            win32gui.EnumWindows(visit,None)
    watcher=threading.Thread(target=watch,daemon=True);watcher.start();book=c=control=None
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.MultiThreadedCalculation.Enabled=False
        book=app.Workbooks.Open(str(bookpath),UpdateLinks=0)
        for n,p in list(files.items())+[('FlowScopeProbe',out/'FlowScopeProbe.bas')]:
            c=book.VBProject.VBComponents.Add(1);c.Name=n;c.CodeModule.AddFromString(code(p))
        c.CodeModule.CodePane.Show();control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled:control.Execute();time.sleep(.3)
        if control.Enabled:
            pane=app.VBE.ActiveCodePane;line=int(pane.GetSelection()[0]);record['compile_error']={'module':pane.CodeModule.Name,'line':pane.CodeModule.Lines(max(1,line-2),5)};save()
            raise RuntimeError('VBA compile failed '+repr(record['compile_error']))
        app.VBE.MainWindow.Visible=False;record['compile']='PASS'
        record['loaded_code_sha256']={c.Name:hashlib.sha256(c.CodeModule.Lines(1,c.CodeModule.CountOfLines).encode('utf8')).hexdigest() for c in book.VBProject.VBComponents if c.Type==1}
        call=lambda n,*v:app.Run("'"+book.Name+"'!"+n,*v)
        call('RPX_Setup')
        input_path=REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs/D02A.json';data=json.loads(input_path.read_text(encoding='utf8'))
        record['global_scope_canonical_input_sha256']=sha(input_path);shutil.copyfile(input_path,out/'scope_input_record.json')
        for sheet,ranges in [('要素定義',[('A','C','points',7),('E','H','regions',7),('J','T','faces',7)]),('材料データ',[('A','E','materials',5)])]:
            ws=book.Worksheets(sheet);ws.Range('A7:T1006' if sheet=='要素定義' else 'A5:Q1004').ClearContents()
            for first,last,key,row in ranges:
                vals=tuple(tuple(v) for v in data[key]);ws.Range(f'{first}{row}:{last}{row+len(vals)-1}').Value2=vals
            ws=None
        ws=book.Worksheets('設定')
        for k,v in data['settings'].items():
            rows=ws.Range('B4:D1000').Value2;chosen=next((i for i,r in enumerate(rows,4) if r[2]==k),None)
            if chosen is None:chosen=next(i for i,r in enumerate(rows,4) if r[2] in [None,''] and r[0] in [None,'']);ws.Cells(chosen,4).Value2=k
            ws.Cells(chosen,2).Value2=v
        ws=None;call('RPX_Speed2pctPreset');record['scope_prepare']=call('FlowScopePrepare');assert record['scope_prepare'].startswith('PASS')
        before=call('FlowScopeKey');record['global_scope_before_sha256']=hashlib.sha256(before.encode('utf8')).hexdigest();save()
        outputs=out/'outputs';outputs.mkdir()
        for item in fixture_manifest['fixtures']:
            source=fixture_root/item['input'];assert sha(source)==item['input_sha256']
            dest=outputs/(item['name']+'.bin');started=time.perf_counter();result=call('FlowMeshFileProbe',str(source),str(dest))
            after=call('FlowScopeKey');record['fixtures'][item['name']]={'result':result,'wall_seconds':time.perf_counter()-started,'input_sha256':sha(source),'output_sha256':sha(dest) if dest.exists() else None,'global_state_unchanged':before==after};save()
            assert result.startswith('PASS'),result
            assert before==after,'Pure-array candidate changed global model/Fs/root/field/program state'
        call('FlowScopeEnd');book.Save();record['status']='VBA_GEOMETRY_PASS';save()
    except BaseException as e:record.update(status='FAIL',error=repr(e));save();raise
    finally:
        try:
            if book is not None:book.Close(False)
        finally:app.Quit()
        stop.set();watcher.join(2);book=c=control=app=None;gc.collect()
        for _ in range(200):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        record['owned_excel_exited']=not psutil.pid_exists(pid);save()
    sys.path.insert(0,str(FLOW/'tests'))
    import compare_excel
    record['python_comparison']=compare_excel.compare(fixture_root,out/'outputs')
    assert sha(original)==BOOK_SHA and sha(FLOW/'SOURCE_MANIFEST.json')==FLOW_SHA
    record['status']='PASS';record['files']={str(p.relative_to(out)):{'bytes':p.stat().st_size,'sha256':sha(p)} for p in out.rglob('*') if p.is_file() and p.name!='summary.json'};save()
    print(json.dumps({'status':'PASS','output':str(out),'owned_excel_exited':record['owned_excel_exited'],'comparison':record['python_comparison']}),flush=True)

if __name__=='__main__':main()
