from native_common import *
from inspect_book import cells
RUNTIME_SCOPE='NATIVE_CONTROL_FLOW_WITH_TEST_ONLY_MOCK_EVALUATE_AND_BAND_NOT_PHYSICAL_SOLVE'
work=R/'tests'/f'ControlledMesh_{time.time_ns()}.xlsm';assert not work.exists();shutil.copy2(source,work)
app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;book=None;rows=[]
internal='''
Public Sub ProbeWeights(ByVal modelPath As String, ByVal velPath As String, ByVal outPath As String)
    Dim v() As Double, w() As Double, f As Integer, i As Long
    MeshLoad modelPath: ReDim v(0 To 12 * re - 1)
    f = FreeFile: Open velPath For Binary Access Read As #f
    For i = 0 To UBound(v): Get #f, , v(i): Next i
    Close #f: ComputeKinematicRefineWeight v, w
    f = FreeFile: Open outPath For Binary Access Write As #f
    For i = 0 To re - 1: Put #f, , w(i): Next i
    Close #f
End Sub
Public Function ProbeGuard() As String
    On Error GoTo Failed
    XCancel = False: RPX_InputBegin
    ThisWorkbook.Worksheets("設定").Cells(48, 2).Value2 = "LEGACY"
    RPX_AssertInputs "policy_changed"
    ProbeGuard = "UNEXPECTED_SUCCESS": RPX_InputEnd: Exit Function
Failed:
    ProbeGuard = "ERROR " & Err.Number & ":" & Err.Description
    Call RPX_InputEnd: XCancel = False
End Function
'''
mock='''Private Sub Evaluate(ByVal mode As String, ByVal label As String)
    Dim eta() As Double, field() As Double, e As Long, k As Long, lower As Double, upper As Double
    meshStage = "evaluate": currentPairValid = False
    RPX_TestHook label
    RPX_AssertInputs "mock_evaluate"
    RFsActive = True: RPX_TotalReferenceLoads
    ReDim RUpperField(0 To 12 * re - 1): ReDim field(0 To 0): ReDim eta(0 To re - 1)
    For e = 0 To re - 1
        eta(e) = 1#
        For k = 0 To 2: RUpperField(12 * e + 2 * k) = RXY(RTri(e, k), 0): Next k
        RUpperField(12 * e + 6) = (RXY(RTri(e, 0), 0) + RXY(RTri(e, 1), 0)) / 2#
        RUpperField(12 * e + 8) = (RXY(RTri(e, 1), 0) + RXY(RTri(e, 2), 0)) / 2#
        RUpperField(12 * e + 10) = (RXY(RTri(e, 2), 0) + RXY(RTri(e, 0), 0)) / 2#
    Next e
    lower = 0.6 + re / 100000#: upper = 0.9 - re / 100000#
    SaveState RCurrent, field, eta, lower, lower, upper, label
    CloneState RBestLower, RCurrent: RBestLower.value = lower
    CloneState RBestUpper, RCurrent: RBestUpper.value = upper
    CloneState RBestPair, RCurrent
    currentPairValid = True: RFsActive = False
End Sub'''
final='''Private Sub RestoreFinalBounds(ByVal mode As String)
    RLowerFs = RBestLower.value: RUpperFs = RBestUpper.value
End Sub'''
runner='''
Public Function ProbeAdapt(ByVal stage As String, ByVal number As Long, ByVal message As String) As String
    Dim e As Long
    On Error GoTo Failed
    XCancel = False: RPX_TestDisarm: RPX_InputEnd: RPX_ReadInputs
    RPX_InputBegin
    If Len(stage) > 0 Then RPX_TestArm stage, number, message
    RPX_RunAdapt "Fs"
    ProbeAdapt = "PASS;elements=" & re & ";hits=" & RPX_TestHits & ";status=" & adaptStatus & ";phase=" & RPX_LoadPhase & ";complete=" & currentPairValid
    For e = 0 To re - 1
        If RBody0(e, 1) <> 0# Or RBody1(e, 1) <> -19# Then Err.Raise 5, , "RESTORED_LOAD_MIXTURE"
    Next e
    Call RPX_InputEnd: Call RPX_TestDisarm: Exit Function
Failed:
    ProbeAdapt = "ERROR " & Err.Number & ":" & Err.Description & ";hits=" & RPX_TestHits
    Call RPX_InputEnd: Call RPX_TestDisarm: XCancel = False
End Function
'''
def replaceproc(text,header,end,body):
    a=text.index(header);b=text.index(end,a)+len(end);return text[:a]+body+text[b:]
try:
    book=app.Workbooks.Open(str(work),UpdateLinks=0);app.EnableEvents=False;inject(book)
    setvalue(book,'ADAPT_MESH_POLICY','LEGACY_THEN_REFINE');setvalue(book,'ADAPT_REFINE_CAP',1000);setvalue(book,'FS_BRACKET_POLICY','AUDITED_UPPER_LIMIT');setvalue(book,'TEST_INJECTION',1)
    cm=book.VBProject.VBComponents('RPX_Adapt').CodeModule;cm.AddFromString(internal)
    compile_book(app,book)
    v=np.load(R/'results/staged_1000_selection.npz')['velocity'].reshape(781,12)
    (R/'results/velocity.bin').write_bytes(v.astype('<f8').tobytes())
    app.Run(f"'{work.name}'!ProbeWeights",str(R/'results/base_input.bin'),str(R/'results/velocity.bin'),str(R/'results/native_scores.bin'))
    w=np.fromfile(R/'results/native_scores.bin',dtype='<f8');py=np.load(R/'results/staged_1000_selection.npz')['score']
    assert np.allclose(w,py,rtol=1e-11,atol=1e-13)
    assert np.array_equal(np.argsort(-w,kind='stable')[:60],np.argsort(-py,kind='stable')[:60])
    rows.append({'test':'native_kinematic_indicator','max_absolute_difference':float(np.max(abs(w-py))),'top_60_ranking_identical':True,'scope':'NATIVE_VBA_INDICATOR_VS_PYTHON'})
    out=app.Run(f"'{work.name}'!ProbeGuard");assert 'INPUT_MISMATCH' in out,out
    assert book.Worksheets('設定').Cells(48,2).Value2=='LEGACY'
    rows.append({'test':'policy_input_changed_events_disabled','result':out,'user_changed_value_preserved':True})
    setvalue(book,'ADAPT_MESH_POLICY','LEGACY_THEN_REFINE');setvalue(book,'ADAPT_REFINE_CAP',1000);setvalue(book,'FS_BRACKET_POLICY','AUDITED_UPPER_LIMIT')
    text=cm.Lines(1,cm.CountOfLines)
    text=replaceproc(text,'Private Sub Evaluate(','End Sub',mock)
    text=replaceproc(text,'Private Sub RestoreFinalBounds(','End Sub',final)
    text=replaceproc(text,'Private Sub RPX_BuildBandMesh(','End Sub','Private Sub RPX_BuildBandMesh(ByRef baseline As RPX_MeshState)\n    RPX_RestoreState baseline\nEnd Sub')
    cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(text+'\r\n'+runner);compile_book(app,book)
    cases=[('normal','',0,'','PASS;'),('known_candidate','staged_refine',5,'HSD_NOT_CONVERGED','PASS;'),
           ('cancel','staged_refine',18,'ANALYSIS_CANCELLED','ERROR 18:'),
           ('oom','staged_refine',7,'OUT_OF_MEMORY','ERROR 7:'),
           ('subscript','staged_refine',9,'SUBSCRIPT','ERROR 9:'),
           ('unknown','staged_refine',5,'UNKNOWN_FAILURE','ERROR 5:'),
           ('audit','staged_refine',5,'PHYSICAL_AUDIT_FAILED','ERROR 5:'),
           ('resource','staged_refine',5,'HSD_STABLE_UNAVAILABLE_MEMORY','ERROR 5:'),
           ('resource_final','final',5,'HSD_STABLE_UNAVAILABLE_MEMORY','ERROR 5:'),
           ('known_refine','refine',5,'NONPOSITIVE_TRIANGLE','PASS;')]
    for name,stage,number,msg,expected in cases:
        out=app.Run(f"'{work.name}'!ProbeAdapt",stage,number,msg)
        assert out.startswith(expected),(name,out)
        if stage:assert ';hits=1' in out,out
        if name=='normal':assert int(out.split('elements=')[1].split(';')[0])<=1000
        if name=='known_candidate':assert 'status=staged_candidate_rejected' in out,out
        rows.append({'test':name,'result':out,'scope':RUNTIME_SCOPE});print(rows[-1],flush=True)
    (R/'results/native_controlled.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
finally:
    if book is not None:book.Close(False)
    app.Quit()
