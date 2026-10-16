"""Build an isolated, opt-in same-Fs upper LP/two-QP recovery sibling."""
from pathlib import Path
import argparse, hashlib, json

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[1]
BASE = REPO / 'versions/G1R19_TerminalRoot_20261016/src/vba/RPX_AccuracyNative.bas'

BRANCH = '''    If mode = "upper" Then
        If UpperRecenterEnabled() Then
            RPX_AccuracyRecoverPrimal = RecoverUpperRecenter(factor, recovered)
            Exit Function
        End If
    End If
'''

DIRECT_BRANCH = '''UpperRecoveryFallback:
    If Not accepted And useUpperRecenter Then
        recenterInvoked = True: recenterSolveId = P4Solve: recenterCandidateId = RPX_HSDCandidateId
        Call RPX_AccuracyGuideReset
        If RecoverUpperRecenterProgram(original, candidate, hasOriginalCenter, recentered) Then
            candidate = recentered
            ok = RPX_AccuracyPrimalAudit(original, candidate, objective)
            If ok Then ok = RPX_FieldGate("upper", candidate, factor, 8, objective, measured, audit, raw)
            If ok Then
                If Not RPX_SchurFinite(measured) Or Not RPX_SchurFinite(audit) Then ok = False
                If measured > 0.999999 Then ok = False
            End If
            UpperRecenterAssert original, recenterSolveId, recenterCandidateId, "direct_upper_recenter_commit"
            If ok Then kind = "original_primal": accepted = True
        End If
    End If
'''

HELPERS = '''
' Opt-in candidate generation only. The legacy path and all original gates stay intact.
Private Function UpperRecenterEnabled() As Boolean
    Dim enabled As Double
    enabled = CDbl(RPX_Setting("ACCURACY_DLL_UPPER_RECENTER", 0))
    If Not RPX_SchurFinite(enabled) Or (enabled <> 0# And enabled <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_UPPER_RECENTER"
    UpperRecenterEnabled = (enabled = 1#)
End Function
Private Function UpperWorkRow(ByRef original As RPX_AccuracyProgram, ByRef workRow As Long) As Boolean
    Dim i As Long, count As Long
    workRow = -1
    If original.mode <> "upper" Or original.q <> 8 Or original.phase <> "reference_total" Then Exit Function
    If original.n <= 0 Or original.equalityCount <= 0 Or original.m <= original.equalityCount Then Exit Function
    For i = 0 To original.equalityCount - 1
        If Not RPX_SchurFinite(original.rhs(i)) Then Exit Function
        If original.rhs(i) <> 0# Then count = count + 1: workRow = i
    Next i
    If count <> 1 Then Exit Function
    For i = original.equalityCount To original.m - 1
        If Not RPX_SchurFinite(original.rhs(i)) Or original.rhs(i) <> 0# Then Exit Function
    Next i
    UpperWorkRow = True
End Function
Private Sub UpperRecenterAssert(ByRef original As RPX_AccuracyProgram, ByVal solveId As Long, ByVal candidateId As Long, ByVal stage As String)
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    RPX_AssertInputs stage
    If Not CurrentProgram(original) Or P4Solve <> solveId Or RPX_HSDCandidateId <> candidateId Then err.Raise 5, "RPX_AccuracyNative", "UPPER_RECENTER_CONTEXT_CHANGED"
End Sub
Private Function NormalizeUpperWork(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double, ByRef workRatio As Double) As Boolean
    Dim workRow As Long, j As Long, i As Long, total As Double, correction As Double, normalized() As Double
    On Error GoTo Rejected
    workRatio = 0#
    If Not UpperWorkRow(original, workRow) Then Exit Function
    If LBound(candidate) <> 0 Or UBound(candidate) <> original.n - 1 Then Exit Function
    For i = 0 To original.n - 1: If Not RPX_SchurFinite(candidate(i)) Then Exit Function
    Next i
    For j = original.rowPtr(workRow) To original.rowPtr(workRow + 1) - 1
        AddTerm total, correction, original.values(j) * candidate(original.colId(j))
    Next j
    workRatio = total / original.rhs(workRow)
    If Not RPX_SchurFinite(workRatio) Or workRatio <= 0# Then Exit Function
    normalized = candidate
    For i = 0 To original.n - 1
        normalized(i) = candidate(i) / workRatio
        If Not RPX_SchurFinite(normalized(i)) Then Exit Function
    Next i
    candidate = normalized: NormalizeUpperWork = True
    Exit Function
Rejected:
    If err.number <> 6 And err.number <> 9 And Left$(err.description, 18) <> "ACCURACY_NONFINITE" Then err.Raise err.number, err.source, err.description
    err.Clear
End Function
Private Function AuditUpperRecenter(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double, ByVal solveId As Long, ByVal candidateId As Long, ByVal stage As String, ByRef workRatio As Double) As Boolean
    Dim objective As Double, raw() As Double, accepted As Boolean
    UpperRecenterAssert original, solveId, candidateId, stage & "_before_audit"
    If NormalizeUpperWork(original, candidate, workRatio) Then
        accepted = RPX_AccuracyPrimalAudit(original, candidate, objective)
        If accepted Then accepted = RPX_FieldGate("upper", candidate, original.factor, 8, objective, AccuracyNativeValue, AccuracyNativeAudit, raw)
        If accepted Then
            If Not RPX_SchurFinite(AccuracyNativeValue) Or Not RPX_SchurFinite(AccuracyNativeAudit) Then accepted = False
            If AccuracyNativeValue > 0.999999 Then accepted = False
        End If
    End If
    UpperRecenterAssert original, solveId, candidateId, stage & "_after_audit"
    RPX_DiagEvent "upper_recenter_candidate", "stage=" & stage & ";Fs=" & original.factor & ";work_ratio=" & workRatio & ";accepted=" & accepted & ";raw_eq=" & AccuracyNativeEquality & ";raw_cone=" & AccuracyNativeCone & ";value=" & AccuracyNativeValue & ";modified_dual_used=0;root_created=0"
    AuditUpperRecenter = accepted
End Function
Private Function RecoverUpperRecenter(ByVal factor As Double, ByRef recovered() As Double) As Boolean
    Dim original As RPX_AccuracyProgram, blankCenter() As Double
    Erase recovered
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    RPX_AssertInputs "upper_recenter_export"
    If Not RPX_AccuracyExport(factor, "upper", original) Then Exit Function
    RecoverUpperRecenter = RecoverUpperRecenterProgram(original, blankCenter, False, recovered)
End Function
Private Function UpperFiniteCenter(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double) As Boolean
    Dim i As Long
    On Error GoTo Rejected
    If LBound(candidate) <> 0 Or UBound(candidate) <> original.n - 1 Then Exit Function
    For i = 0 To original.n - 1: If Not RPX_SchurFinite(candidate(i)) Then Exit Function
    Next i
    UpperFiniteCenter = True: Exit Function
Rejected:
    If err.number <> 9 Then err.Raise err.number, err.source, err.description
    err.Clear
End Function
Private Function RecoverUpperRecenterProgram(ByRef original As RPX_AccuracyProgram, ByRef initialCenter() As Double, ByVal hasInitialCenter As Boolean, ByRef recovered() As Double) As Boolean
    Dim qp As RPX_AccuracyProgram
    Dim center() As Double, candidate() As Double, modifiedDual() As Double
    Dim alpha As Double, lpSeconds As Double, qpSeconds As Double, workRatio As Double, factor As Double
    Dim i As Long, trialIndex As Long, workRow As Long, solveId As Long, candidateId As Long
    Dim ok As Boolean, finiteCandidate As Boolean, oldRejected As Boolean, auditSaved As Boolean
    Dim auditState As RPX_AuditSnapshot, number As Long, message As String, errorSource As String
    Erase recovered
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    RPX_AssertInputs "upper_recenter_program"
    If Not CurrentProgram(original) Then err.Raise 5, "RPX_AccuracyNative", "UPPER_RECENTER_CONTEXT_CHANGED"
    If Not UpperWorkRow(original, workRow) Then AccuracyNativeReason = "UPPER_RECENTER_NONHOMOGENEOUS_SCOPE_REJECTED": Exit Function
    factor = original.factor
    solveId = P4Solve: candidateId = RPX_HSDCandidateId
    alpha = CDbl(RPX_Setting("ACCURACY_DLL_ALPHA", 0.01))
    lpSeconds = CDbl(RPX_Setting("ACCURACY_DLL_UPPER_LP_SECONDS", 30))
    qpSeconds = CDbl(RPX_Setting("ACCURACY_DLL_SECONDS", 120))
    If Not RPX_SchurFinite(alpha) Or alpha <= 0# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_ALPHA"
    If Not RPX_SchurFinite(lpSeconds) Or lpSeconds <= 0# Or lpSeconds > 600# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_UPPER_RECENTER_LP_SECONDS"
    If Not RPX_SchurFinite(qpSeconds) Or qpSeconds <= 0# Or qpSeconds > 600# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_UPPER_RECENTER_QP_SECONDS"
    RPX_DiagAuditSave auditState: oldRejected = RPX_AuditRejected: auditSaved = True
    On Error GoTo Failed
    If hasInitialCenter Then
        If UpperFiniteCenter(original, initialCenter) Then
            candidate = initialCenter: finiteCandidate = True
            RPX_DiagEvent "upper_recenter_center", "source=reused_direct_original_lp;same_exported_program=1;Fs=" & factor
        End If
    End If
    If Not finiteCandidate Then
        RPX_DiagEvent "upper_recenter_center", "source=fresh_original_lp;reason=direct_center_absent_or_invalid;Fs=" & factor
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_lp_before"
        RPX_TestHook "upper_recenter_lp_before"
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_lp_start"
        finiteCandidate = RPX_AccuracySolve(original, lpSeconds, candidate, modifiedDual)
        Erase modifiedDual ' Neither this LP dual nor a modified QP dual is committed here.
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_lp_after"
    End If
    If Not finiteCandidate Then GoTo Restore
    ok = AuditUpperRecenter(original, candidate, solveId, candidateId, "original_lp", workRatio)
    If ok Then GoTo Restore
    center = candidate
    For trialIndex = 1 To 2
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_qp_before"
        RPX_TestHook "upper_recenter_qp_before"
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_qp_start"
        qp = original: AddBound qp, -1, 0.999998
        For i = 0 To original.n - 1
            qp.pdiag(i) = alpha: qp.cost(i) = -alpha * center(i)
            If Not RPX_SchurFinite(qp.cost(i)) Then err.Raise 5, "RPX_AccuracyNative", "UPPER_RECENTER_NONFINITE_QP_CENTER"
        Next i
        ' Literal original rows plus the cap: Python recipe uses no cone row scaling.
        finiteCandidate = RPX_AccuracySolve(qp, qpSeconds, candidate, modifiedDual)
        Erase modifiedDual ' An auxiliary cap/QP dual is NEVER an original root guide.
        UpperRecenterAssert original, solveId, candidateId, "upper_recenter_qp_after"
        If Not finiteCandidate Then GoTo Restore
        ok = AuditUpperRecenter(original, candidate, solveId, candidateId, "cap_qp_" & trialIndex, workRatio)
        If ok Then Exit For
        center = candidate ' Same Fs and the normalized rejected candidate only; never an old field.
    Next trialIndex
Restore:
    Erase modifiedDual
    If auditSaved Then
        RPX_DiagAuditRestore auditState: RPX_AuditRejected = oldRejected
    End If
    On Error GoTo 0
    If number <> 0 Then err.Raise number, errorSource, message
    UpperRecenterAssert original, solveId, candidateId, "upper_recenter_return"
    If ok Then
        recovered = candidate: AccuracyNativeReason = "ORIGINAL_PHYSICAL_PRIMAL_WITNESS"
        CaptureAccuracyProof original, candidate, False
    Else
        Erase recovered: AccuracyNativeReason = "ORIGINAL_PHYSICAL_PRIMAL_REJECTED"
    End If
    RPX_DiagEvent "upper_recenter_complete", "accepted=" & ok & ";Fs=" & factor & ";q=8;max_qp_trials=2;lp_seconds=" & lpSeconds & ";qp_seconds_each=" & qpSeconds & ";modified_dual_used=0;root_created=0"
    RecoverUpperRecenterProgram = ok: Exit Function
Failed:
    number = err.number: message = err.description: errorSource = err.source: ok = False
    Resume Restore
End Function
'''


def norm(text):
    return text.replace('\r\n', '\n').replace('\r', '\n')


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def direct_changes():
    return [
        ('    Dim normalized() As Double, normalizedObjective As Double, normalizedValue As Double\n',
         '    Dim normalized() As Double, normalizedObjective As Double, normalizedValue As Double\n'
         '    Dim recentered() As Double, useUpperRecenter As Boolean, hasOriginalCenter As Boolean, recenterInvoked As Boolean\n'
         '    Dim recenterSolveId As Long, recenterCandidateId As Long\n'),
        ('    If Not RPX_AccuracyExport(factor, mode, original) Then Exit Function\n',
         '    If Not RPX_AccuracyExport(factor, mode, original) Then Exit Function\n'
         '    If mode = "upper" Then useUpperRecenter = UpperRecenterEnabled()\n'),
        ('    If Not RPX_AccuracySolve(original, seconds, candidate, originalDual) Then GoTo Restore\n',
         '    hasOriginalCenter = RPX_AccuracySolve(original, seconds, candidate, originalDual)\n'
         '    If Not hasOriginalCenter Then GoTo UpperRecoveryFallback\n'),
        ('Restore:\n', DIRECT_BRANCH + 'Restore:\n'),
        ('    On Error GoTo 0\n    RPX_DiagEvent "direct_dll_end",',
         '    On Error GoTo 0\n'
         '    If recenterInvoked And number <> 0 Then err.Raise number, "RPX_AccuracyNative", message\n'
         '    RPX_DiagEvent "direct_dll_end",'),
    ]


def expected():
    text = norm(BASE.read_text(encoding='utf8'))
    first = text.index('Public Function RPX_AccuracyRecoverPrimal(')
    end = text.index('\nEnd Function', first)
    anchor = '    Erase recovered\n'
    old = text[first:end]
    if old.count(anchor) != 1:
        raise ValueError('Legacy recovery anchor changed')
    new = old.replace(anchor, BRANCH + anchor, 1)
    text = text[:first] + new + text[end:]
    first = text.index('Public Function RPX_AccuracyDirectBound(')
    end = text.index('\nEnd Function', first)
    direct = text[first:end]
    for before, after in direct_changes():
        if direct.count(before) != 1:
            raise ValueError('Direct recovery anchor changed: ' + before)
        direct = direct.replace(before, after, 1)
    return text[:first] + direct + text[end:] + HELPERS


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--initialize', action='store_true')
    args = parser.parse_args()
    text = expected()
    source = ROOT / 'src/vba/RPX_AccuracyNative.bas'
    imported = ROOT / 'import/RPX_AccuracyNative.bas'
    if args.initialize:
        for path, encoding in [(source, 'utf8'), (imported, 'cp932')]:
            path.parent.mkdir(parents=True, exist_ok=True)
            if not path.exists():
                path.write_bytes(text.replace('\n', '\r\n').encode(encoding))
    if norm(source.read_text(encoding='utf8')) != text:
        raise ValueError('Candidate differs from controlled opt-in patch')
    if source.read_text(encoding='utf8') != imported.read_text(encoding='cp932'):
        raise ValueError('UTF8/CP932 mismatch')
    tracked = ['build_overlay.py', 'README_JA.md', 'SETTINGS_CONTRACT.json',
               'tests/test_contract.py', 'tests/normalization_reference.py',
               'tests/UpperRecoveryArrayProbe.bas', 'tests/MICROTEST_JA.md']
    manifest = dict(schema='rpfem.upper_recovery.same_fs.v1',
                    status='ISOLATED_OPT_IN_PROTOTYPE; EXCEL_RUNTIME_PENDING',
                    production_artifacts_modified=False, solver_calls=0,
                    baseline=BASE.relative_to(REPO).as_posix(), baseline_sha256=sha(BASE),
                    baseline_manifest='versions/G1R19_TerminalRoot_20261016/SOURCE_MANIFEST.json',
                    baseline_manifest_sha256=sha(BASE.parents[2] / 'SOURCE_MANIFEST.json'),
                    source='src/vba/RPX_AccuracyNative.bas', source_sha256=sha(source),
                    imported='import/RPX_AccuracyNative.bas', import_sha256=sha(imported),
                    profile=json.loads((ROOT / 'SETTINGS_CONTRACT.json').read_text(encoding='utf8')),
                    files={str(path): sha(ROOT / path) for path in tracked})
    (ROOT / 'SOURCE_MANIFEST.json').write_text(json.dumps(manifest, indent=2, ensure_ascii=False) + '\n', encoding='utf8')
    print(json.dumps(dict(status=manifest['status'], source_sha256=manifest['source_sha256'],
                          manifest_sha256=sha(ROOT / 'SOURCE_MANIFEST.json'))))


if __name__ == '__main__':
    main()
