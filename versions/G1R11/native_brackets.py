from native_common import *
def replaceproc(s,header,end,body):
    a=s.index(header);b=s.index(end,a)+len(end);return s[:a]+body+s[b:]
mock='''Private Function FsValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double) As Double
    If mode = "lower" Then
        mockCalls = mockCalls + 1
        If factor > mockMax Then mockMax = factor
        If mockErrorNumber <> 0 And factor = mockErrorAt Then
            Dim n As Long, msg As String
            n = mockErrorNumber: msg = mockErrorText: mockErrorNumber = 0
            If n = 18 Then XCancel = True
            err.Raise n, "BracketTestOnly", msg
        End If
        RValue = mockLower / factor
    Else
        RValue = 0.77 / factor
    End If
    XStatus = "OPTIMAL": RPX_AuditRejected = False
    ReDim xx(0 To 0): xx(0) = RValue
    XObjective = RValue: If mode = "lower" Then XObjective = -RValue
    FsValue = RValue - target
End Function'''
runner='''
Public Function ProbeBracket(ByVal kind As String) As String
    Dim upper As Double, lower As Double, cap As Boolean
    On Error GoTo Failed
    If kind = "legacy_prefix" Then
        MeshLoad "BASE_INPUT"
    Else
        MeshLoad "MODEL_INPUT"
    End If
    XCancel = False: RFsActive = True: RPX_TotalReferenceLoads
    Call RPX_FsBracketReset: RFsTolerance = 0.0001: RSubdivisions = 4
    RFsSearchIncomplete = False: RFsLargestRootBracket = 0#: mockLower = 0.66
    upper = RPX_FsEndpoint("upper", 0.77)
    mockCalls = 0: mockMax = 0#: mockErrorNumber = 0
    Select Case kind
        Case "positive_sign": mockLower = 0.9
        Case "numeric": mockErrorNumber = 5: mockErrorText = "HSD_NOT_CONVERGED"
        Case "oom": mockErrorNumber = 7: mockErrorText = "OUT_OF_MEMORY"
        Case "subscript": mockErrorNumber = 9: mockErrorText = "SUBSCRIPT"
        Case "unknown": mockErrorNumber = 5: mockErrorText = "UNKNOWN_ERROR"
        Case "audit": mockErrorNumber = 5: mockErrorText = "PHYSICAL_AUDIT_FAILED"
        Case "resource": mockErrorNumber = 5: mockErrorText = "HSD_STABLE_UNAVAILABLE_MEMORY"
        Case "cancel": mockErrorNumber = 18: mockErrorText = "ANALYSIS_CANCELLED"
        Case "quadrature": RSubdivisions = 8
        Case "model": RMaterials(0).cohesion = RMaterials(0).cohesion + 1#
        Case "rigid": RR = 1
        Case "wrong_endpoint": upper = upper + 0.01
        Case "uncertified": RPX_AuditRejected = True
    End Select
    mockErrorAt = upper
    cap = UpperLimitUsable(upper, FsMeshKey())
    lower = RPX_FsEndpoint("lower", 0#, 0.1, upper)
    ProbeBracket = "PASS;lower=" & lower & ";upper=" & upper & ";cap_eligible=" & cap & ";calls=" & mockCalls & ";max_trial=" & mockMax & ";a=" & lowerBracket.a & ";b=" & lowerBracket.b & ";fa=" & lowerBracket.fa & ";fb=" & lowerBracket.fb
    Exit Function
Failed:
    ProbeBracket = "ERROR " & err.number & ":" & err.description & ";calls=" & mockCalls
    XCancel = False
End Function
'''
work=R/'tests'/f'Brackets_{time.time_ns()}.xlsm';shutil.copy2(source,work)
app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;book=None;rows=[]
try:
    book=app.Workbooks.Open(str(work),UpdateLinks=0);app.EnableEvents=False;inject(book)
    setvalue(book,'FS_BRACKET_POLICY','AUDITED_UPPER_LIMIT');setvalue(book,'ADAPT_MESH_POLICY','LEGACY_THEN_REFINE')
    cm=book.VBProject.VBComponents('RPX_Fs').CodeModule;s=cm.Lines(1,cm.CountOfLines)
    s=s.replace('Option Explicit','Option Explicit\nPrivate mockLower As Double, mockMax As Double, mockCalls As Long, mockErrorNumber As Long, mockErrorAt As Double, mockErrorText As String',1)
    s=replaceproc(s,'Private Function FsValue(','End Function',mock)
    s+=runner.replace('MODEL_INPUT',str(R/'results/native1000_input.bin')).replace('BASE_INPUT',str(R/'results/base_input.bin'))
    cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(s)
    # Endpoint audits are replaced ONLY in this scheduler test workbook.
    cm=book.VBProject.VBComponents('RPX_Audit').CodeModule;s=cm.Lines(1,cm.CountOfLines)
    s=replaceproc(s,'Public Sub RPX_AuditLower(','End Sub','Public Sub RPX_AuditLower(Optional ByVal forGate As Boolean = False)\n    RValue = xx(0)\nEnd Sub')
    s=replaceproc(s,'Public Sub RPX_AuditUpper(','End Sub','Public Sub RPX_AuditUpper(Optional ByVal forGate As Boolean = False)\n    RValue = xx(0)\nEnd Sub')
    cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(s);compile_book(app,book)
    for kind in ['legacy_prefix','normal','numeric','positive_sign','quadrature','model','wrong_endpoint','uncertified','oom','subscript','unknown','audit','resource','cancel']:
        out=app.Run(f"'{work.name}'!ProbeBracket",kind)
        expected={'oom':'ERROR 7:','subscript':'ERROR 9:','unknown':'ERROR 5:UNKNOWN_ERROR','audit':'ERROR 5:PHYSICAL_AUDIT_FAILED','resource':'ERROR 5:HSD_STABLE_UNAVAILABLE_MEMORY','cancel':'ERROR 18:'}.get(kind,'PASS;')
        assert out.startswith(expected),(kind,out)
        if kind in ['legacy_prefix','quadrature','model','wrong_endpoint','uncertified']:assert 'cap_eligible=False' in out,out
        if kind=='normal':assert 'cap_eligible=True' in out and float(out.split('max_trial=')[1].split(';')[0])<.78,out
        if kind=='numeric':assert float(out.split('max_trial=')[1].split(';')[0])>=2 and int(out.split('calls=')[1].split(';')[0])>0,out
        if kind=='positive_sign':assert float(out.split('lower=')[1].split(';')[0])>.89,out
        rows.append({'test':kind,'result':out,'scope':'NATIVE_BRACKET_CONTROL_WITH_MOCK_VALUE_AND_AUDITS_NOT_PHYSICAL_SOLVE'});print(rows[-1],flush=True)
    setvalue(book,'FS_BRACKET_POLICY','LEGACY')
    out=app.Run(f"'{work.name}'!ProbeBracket",'normal');assert out.startswith('PASS;') and 'cap_eligible=False' in out,out
    rows.append({'test':'legacy_scheduler','result':out,'scope':'MOCK_CALLBACK'})
    (R/'results/native_brackets.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
finally:
    if book is not None:book.Close(False)
    app.Quit()
