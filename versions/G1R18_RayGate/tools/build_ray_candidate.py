"""Reproducible original-dual guide variant; G1R18 production is retained."""
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT.parent / 'G1R17' / 'src/vba'

def once(source, old, new):
    if source.count(old) != 1:
        raise ValueError(f'Expected one occurrence: {old[:100]!r}')
    return source.replace(old, new, 1)

def write(name, source):
    for folder, encoding in [('src/vba', 'utf-8'), ('import', 'cp932')]:
        path = ROOT / folder / (name + '.bas')
        path.parent.mkdir(parents=True, exist_ok=True)
        path.write_bytes(source.replace('\n', '\r\n').encode(encoding))

def main():
    source = (BASE / 'RPX_Conic.bas').read_text(encoding='utf-8')
    declarations = (ROOT / 'tools/original_dual_declarations.bas.txt').read_text(encoding='utf-8')
    source = once(source, 'Public OriginalEqualityError As Double, OriginalConeError As Double, OriginalCandidateObjective As Double', declarations)
    source = once(source, 'Public Sub RPX_Begin(ByVal variables As Long)\n    RPX_NativeBulkReset', 'Public Sub RPX_Begin(ByVal variables As Long)\n    RPX_ClearOriginalCompiled\n    RPX_NativeBulkReset')
    source = once(source, '    For i = 0 To XN - 1: XQ(i) = XQ(i) / XQScale: Next i', '''    RPX_ClearOriginalCompiled
    XOriginalCompiled.n = XN: XOriginalCompiled.m = XM: XOriginalCompiled.s = XS: XOriginalCompiled.b = XB
    XOriginalCompiled.cost = XQ
    XOriginalCompiled.modeName = RPX_OrderMode: XOriginalCompiled.subdivisions = RSubdivisions
    XOriginalCompiled.inputEpoch = RPX_InputEpoch: XOriginalCompiled.affine = AFActive
    XOriginalCompiled.loadPhase = RPX_LoadPhase
    If rn > 0 Then XOriginalCompiled.modelKey = RPX_ExactModel()
    For i = 0 To XN - 1: XQ(i) = XQ(i) / XQScale: Next i''')
    source = once(source, '    ReDim XEId(0 To XEPtr(XM)): ReDim XEVal(0 To XEPtr(XM))', '    ReDim XEId(0 To XEPtr(XM)): ReDim XEVal(0 To XEPtr(XM))\n    ReDim XOriginalCompiled.values(0 To XEPtr(XM)): ReDim XOriginalCompiled.rhs(0 To XM)')
    source = once(source, '        Set row = indexedRows(i + 1): valueScale = 0#', '        Set row = indexedRows(i + 1): valueScale = 0#\n        XOriginalCompiled.rhs(i) = indexedRhs(i + 1)')
    source = once(source, '            XEId(p) = keyIds(ki): XEVal(p) = row(keyIds(ki)) / valueScale: p = p + 1', '            XOriginalCompiled.values(p) = CDbl(row(keyIds(ki)))\n            XEId(p) = keyIds(ki): XEVal(p) = row(keyIds(ki)) / valueScale: p = p + 1')
    source = once(source, '    Set rawEq = Nothing: Set rawCones = Nothing: Set rawRhs = Nothing', '    XOriginalCompiled.ready = True\n    Set rawEq = Nothing: Set rawCones = Nothing: Set rawRhs = Nothing')
    source += '\n' + (ROOT / 'tools/original_dual_gate.bas.txt').read_text(encoding='utf-8')
    source += '\n' + (ROOT / 'tools/original_dual_capture.bas.txt').read_text(encoding='utf-8')
    write('RPX_Conic', source)

    source = (BASE / 'RPX_LowerCache.bas').read_text(encoding='utf-8')
    source = once(source, 'Private savedRowNorm() As Double', 'Private savedRowNorm() As Double\nPrivate savedOriginal As RPX_OriginalCompiledData')
    source = once(source, 'Public Sub RPX_LowerCacheReset()\n    cacheValid', 'Public Sub RPX_LowerCacheReset()\n    Dim blankOriginal As RPX_OriginalCompiledData\n    savedOriginal = blankOriginal\n    cacheValid')
    source = once(source, '    bytes = bytes + 256# * RReactions.count + 256# * XM', '    bytes = bytes + 256# * RReactions.count + 256# * XM\n    bytes = bytes + 8# * (XN + XM + 1# + XEPtr(XM) + 1#) + 2# * Len(XOriginalCompiled.modelKey) + 512#')
    source = once(source, '    savedRowNorm = XEqRowNorm', '    savedRowNorm = XEqRowNorm: savedOriginal = XOriginalCompiled')
    source = once(source, '    XEqRowNorm = savedRowNorm', '    XEqRowNorm = savedRowNorm: XOriginalCompiled = savedOriginal')
    write('RPX_LowerCache', source)

    source = (BASE / 'RPX_Solve.bas').read_text(encoding='utf-8')
    source = once(source, 'Public Sub RPX_ResetPreparedState()\n    RPX_NativeBulkReset', 'Public Sub RPX_ResetPreparedState()\n    RPX_ClearOriginalCompiled\n    RPX_NativeBulkReset')
    write('RPX_Solve', source)

    source = (BASE / 'RPX_HSDCandidate.bas').read_text(encoding='utf-8')
    source = once(source, 'Public Sub RPX_HSDObservationBegin(ByVal mode As String)\n    Dim i As Long', 'Public Sub RPX_HSDObservationBegin(ByVal mode As String)\n    Dim i As Long\n    RPX_OriginalGuideClear')
    source += '\nPublic Function RPX_HSDGeometryEpoch() As Long\n    RPX_HSDGeometryEpoch = geometryEpoch\nEnd Function\n'
    old = '    If Len(candidate.mode) = 0 Or candidate.kind = "PRIMAL_INFEASIBILITY_CERT" Then\n        candidate.physicalGate = True: candidate.valid = True: RPX_HSDValidateCandidate = True: Exit Function\n    End If'
    new = '''    If AFActive And candidate.kind <> "OPTIMAL_POINT" Then Exit Function
    If candidate.kind = "PRIMAL_INFEASIBILITY_CERT" And candidate.mode = "upper" And candidate.subdivisions = 8 Then
        If RPX_OriginalPhysicalZeroCostUpper() Then
            If Not RPX_OriginalFarkasGate(candidate) Then Exit Function
            candidate.outputStatus = "ORIGINAL_FARKAS_GUIDE"
            Erase candidate.x: Erase candidate.field
            candidate.physicalGate = True: candidate.valid = True
            RPX_HSDValidateCandidate = True: Exit Function
        End If
    End If
    If Len(candidate.mode) = 0 Or candidate.kind = "PRIMAL_INFEASIBILITY_CERT" Then
        candidate.physicalGate = True: candidate.valid = True: RPX_HSDValidateCandidate = True: Exit Function
    End If'''
    source = once(source, old, new)
    source = once(source, '    xx = candidate.x: xy = candidate.y: XSlack = candidate.slack: XDual = candidate.dual', '''    If candidate.outputStatus = "ORIGINAL_FARKAS_GUIDE" Then
        Erase xx: Erase XSlack
        xy = candidate.y: XDual = candidate.dual
    Else
        xx = candidate.x: xy = candidate.y: XSlack = candidate.slack: XDual = candidate.dual
    End If''')
    write('RPX_HSDCandidate', source)

    source = (BASE / 'RPX_Fs.bas').read_text(encoding='utf-8')
    source = once(source, '    statusB As String\nEnd Type\nPrivate Type FsTrialPoint', '    statusB As String\n    originalGuideA As Boolean\nEnd Type\nPrivate Type FsTrialPoint')
    source = once(source, '    If XStatus <> "OPTIMAL" Then\n        RProgramKind = ""\n        Exit Function\n    End If', '''    If XStatus <> "OPTIMAL" Then
        If mode <> "upper" Or point.residual < 0# Or Not RPX_OriginalGuideCurrent(factor) Then
            RProgramKind = "": Exit Function
        End If
    End If''')
    source = once(source, '        If RFsActive And mode = "upper" And XStatus = "INFEASIBLE" Then', '        If RFsActive And mode = "upper" And (XStatus = "INFEASIBLE" Or RPX_OriginalGuideCurrent(strengthFactor)) Then')
    source = once(source, '    If XStatus = "INFEASIBLE" Or XStatus = "UNBOUNDED" Then FsValue = 1#', '''    If XStatus = "ORIGINAL_FARKAS_GUIDE" Then
        If mode = "upper" And RPX_OriginalGuideCurrent(factor) Then
            FsValue = 1#
        Else
            XStatus = "NUMERICAL_UNKNOWN"
        End If
    ElseIf XStatus = "INFEASIBLE" Or XStatus = "UNBOUNDED" Then
        FsValue = 1#
    End If''')
    source = once(source, '    If stored.statusA <> "OPTIMAL" Or stored.statusB <> "OPTIMAL" Then Exit Function', '''    If stored.statusB <> "OPTIMAL" Then Exit Function
    If stored.statusA <> "OPTIMAL" Then
        If side <> "upper" Or Not stored.originalGuideA Or stored.statusA <> "ORIGINAL_FARKAS_GUIDE" Or stored.subdivisions <> 8 Then Exit Function
    End If''')
    source = once(source, '    If Not FsArrayReady(fieldA) Or Not FsArrayReady(fieldB) Then Exit Sub', '''    If Not FsArrayReady(fieldB) Then Exit Sub
    If Not FsArrayReady(fieldA) Then
        If mode <> "upper" Or statusA <> "ORIGINAL_FARKAS_GUIDE" Or RSubdivisions <> 8 Then Exit Sub
    End If''')
    source = once(source, '        upperBracket.statusA = statusA: upperBracket.statusB = statusB', '        upperBracket.statusA = statusA: upperBracket.statusB = statusB\n        upperBracket.originalGuideA = (statusA = "ORIGINAL_FARKAS_GUIDE" And RSubdivisions = 8)')
    write('RPX_Fs', source.rstrip() + '\n')
    accuracy = ROOT.parent / 'G1R18' / 'src/vba'
    source = (accuracy / 'RPX_AccuracyMesh.bas').read_text(encoding='utf-8')
    source += '\n' + (ROOT / 'tools/c0_slope_wedge_mesh.bas.txt').read_text(encoding='utf-8')
    write('RPX_AccuracyMesh', source)
    source = (accuracy / 'RPX_Adapt.bas').read_text(encoding='utf-8')
    source = once(source, '    Call ContinueLowerAngularFan(mode)\n    Call RecheckIndependentQuadrature(mode)', '    Call ContinueLowerAngularFan(mode)\n    Call ContinueUpperC0Wedge(mode)\n    Call RecheckIndependentQuadrature(mode)')
    source = once(source, 'Public Sub RPX_RunAdapt(ByVal mode As String)', (ROOT / 'tools/c0_slope_wedge_adapt.bas.txt').read_text(encoding='utf-8') + '\nPublic Sub RPX_RunAdapt(ByVal mode As String)')
    write('RPX_Adapt', source)
    print('Created independent G1R18_RayGate overrides and CP932 imports.')

if __name__ == '__main__':
    main()
