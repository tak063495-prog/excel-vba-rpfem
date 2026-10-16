"""Opt-in original upper dual, one free-y LSMR correction, no DLL rebuild."""
from pathlib import Path
import argparse, hashlib, json

ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[1]
BASE=REPO/'versions/G1R19_UpperRecovery_20261016/src/vba/RPX_AccuracyNative.bas'

HELPERS='''
' The cone dual is immutable. A correction is a guide only after the original audit.
Private Function DualLinearRepairEnabled() As Boolean
    Dim enabled As Double
    enabled = CDbl(RPX_Setting("ACCURACY_DLL_DUAL_LINEAR_REPAIR", 0))
    If Not RPX_SchurFinite(enabled) Or (enabled <> 0# And enabled <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_DUAL_LINEAR_REPAIR"
    DualLinearRepairEnabled = (enabled = 1#)
End Function
Private Sub SetUpperDualRegularization(ByRef dp As RPX_AccuracyProgram)
    Dim coefficient As Double, i As Long
    coefficient = CDbl(RPX_Setting("ACCURACY_DLL_DUAL_REGULARIZATION", 0))
    If Not RPX_SchurFinite(coefficient) Or coefficient < 0# Then err.Raise 5, "RPX_AccuracyNative", "INVALID_DUAL_REGULARIZATION"
    For i = 0 To dp.n - 1: dp.pdiag(i) = coefficient: Next i
End Sub
Private Sub DualRepairAssert(ByRef original As RPX_AccuracyProgram, ByVal solveId As Long, ByVal candidateId As Long, ByVal stage As String)
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    RPX_AssertInputs stage
    If Not CurrentProgram(original) Or P4Solve <> solveId Or RPX_HSDCandidateId <> candidateId Then err.Raise 5, "RPX_AccuracyNative", "DUAL_REPAIR_CONTEXT_CHANGED"
End Sub
Private Sub OriginalEqualityTranspose(ByRef original As RPX_AccuracyProgram, ByRef transposed As RPX_AccuracyProgram)
    Dim counts() As Long, cursor() As Long, i As Long, j As Long, row As Long, p As Long
    transposed.n = original.equalityCount: transposed.m = original.n
    transposed.nnz = original.rowPtr(original.equalityCount)
    ReDim counts(0 To original.n - 1): ReDim cursor(0 To original.n - 1)
    For j = 0 To transposed.nnz - 1: counts(original.colId(j)) = counts(original.colId(j)) + 1: Next j
    ReDim transposed.rowPtr(0 To original.n)
    For i = 0 To original.n - 1
        transposed.rowPtr(i + 1) = transposed.rowPtr(i) + counts(i): cursor(i) = transposed.rowPtr(i)
    Next i
    ReDim transposed.colId(0 To transposed.nnz - 1): ReDim transposed.values(0 To transposed.nnz - 1)
    For i = 0 To original.equalityCount - 1
        For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1
            row = original.colId(j): p = cursor(row): cursor(row) = p + 1
            transposed.colId(p) = i: transposed.values(p) = original.values(j)
        Next j
    Next i
End Sub
Private Function DualConeBitsUnchanged(ByRef source() As Double, ByRef updated() As Double, ByVal first As Long, ByVal count As Long) As Boolean
    Dim i As Long, sourceBits As LongLong, updatedBits As LongLong
    For i = first To count - 1
        If Not RPX_SchurFinite(source(i)) Or Not RPX_SchurFinite(updated(i)) Then Exit Function
        If source(i) <> updated(i) Then Exit Function
        ' Equal finite nonzero doubles have identical bits; distinguish +/-0 explicitly.
        If source(i) = 0# Then
            CopyBytes VarPtr(sourceBits), VarPtr(source(i)), 8
            CopyBytes VarPtr(updatedBits), VarPtr(updated(i)), 8
            If sourceBits <> updatedBits Then Exit Function
        End If
    Next i
    DualConeBitsUnchanged = True
End Function
Private Sub CaptureDualRepairStage(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double, ByVal stage As String, ByVal accepted As Boolean)
    Dim enabled As Double, path As String
    enabled = CDbl(RPX_Setting("ACCURACY_DLL_CAPTURE", 0))
    If Not RPX_SchurFinite(enabled) Or (enabled <> 0# And enabled <> 1#) Then err.Raise 5, "RPX_AccuracyNative", "INVALID_ACCURACY_DLL_CAPTURE"
    If enabled = 0# Then Exit Sub
    proofSequence = proofSequence + 1
    ' Distinct filename: rejected source/corrected candidates are NOT accepted accuracy proofs.
    path = RPX_DiagPath & ".dual_repair_" & stage & "_solve_" & P4Solve & "_" & proofSequence & ".bin"
    RPX_AccuracyWriteProof path, original, candidate
    RPX_DiagEvent "dual_repair_capture", "stage=" & stage & ";path=" & path & ";layout=AC18P001_original_dual;accepted=" & accepted & ";physical_bound=0;candidate_only=1"
End Sub
Private Function OriginalDualArrayValid(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double) As Boolean
    Dim i As Long
    On Error GoTo Rejected
    If LBound(candidate) <> 0 Or UBound(candidate) <> original.m - 1 Then Exit Function
    For i = 0 To original.m - 1: If Not RPX_SchurFinite(candidate(i)) Then Exit Function
    Next i
    OriginalDualArrayValid = True: Exit Function
Rejected:
    If err.number <> 9 Then err.Raise err.number, err.source, err.description
    err.Clear
End Function
Private Function TryOriginalDualLinearRepair(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double) As Boolean
    Dim transposed As RPX_AccuracyProgram, sourceDual() As Double, updated() As Double
    Dim rhs() As Double, compensation() As Double, zeroCorrection() As Double, correction() As Double
    Dim descriptor(0 To 63) As Byte, report(0 To 55) As Byte
    Dim i As Long, j As Long, rc As Long, solveId As Long, candidateId As Long, number As Long
    Dim beforeStat As Double, beforeCone As Double, beforeObjective As Double, started As Double, seconds As Double
    Dim message As String, errorSource As String, accepted As Boolean, unchangedCone As Boolean, alreadyAccepted As Boolean
    If original.mode <> "upper" Then Exit Function
    If Not DualLinearRepairEnabled() Then Exit Function
    If XCancel Then err.Raise 18, "RPX_AccuracyNative", "ANALYSIS_CANCELLED"
    solveId = P4Solve: candidateId = RPX_HSDCandidateId
    DualRepairAssert original, solveId, candidateId, "dual_repair_begin"
    If original.equalityCount <= 0 Or original.rowPtr(original.equalityCount) <= 0 Then Exit Function
    If Not OriginalDualArrayValid(original, candidate) Then Exit Function
    On Error GoTo Failed
    sourceDual = candidate
    alreadyAccepted = RPX_AccuracyDualAudit(original, sourceDual)
    beforeStat = AccuracyNativeDualStat: beforeCone = AccuracyNativeDualCone: beforeObjective = AccuracyNativeDualObjective
    If alreadyAccepted Then Exit Function
    If Not RPX_SchurFinite(beforeStat) Or Not RPX_SchurFinite(beforeCone) Or Not RPX_SchurFinite(beforeObjective) Then Exit Function
    If beforeCone > 0.00000001 Then Exit Function ' y cannot repair a rejected cone dual.
    CaptureDualRepairStage original, sourceDual, "before", False
    OriginalEqualityTranspose original, transposed
    ReDim rhs(0 To original.n - 1): ReDim compensation(0 To original.n - 1)
    ReDim zeroCorrection(0 To original.equalityCount - 1): ReDim correction(0 To original.equalityCount - 1)
    For i = 0 To original.n - 1: rhs(i) = -original.originalCost(i): Next i
    For i = 0 To original.m - 1
        For j = original.rowPtr(i) To original.rowPtr(i + 1) - 1
            AddTerm rhs(original.colId(j)), compensation(original.colId(j)), -original.values(j) * sourceDual(i)
        Next j
    Next i
    PackLong descriptor, 0, 64: PackLong descriptor, 4, 1
    PackLong descriptor, 8, transposed.n: PackLong descriptor, 12, transposed.m: PackLong descriptor, 16, transposed.nnz
    PackPointer descriptor, 24, VarPtr(transposed.rowPtr(0)): PackPointer descriptor, 32, VarPtr(transposed.colId(0))
    PackPointer descriptor, 40, VarPtr(transposed.values(0)): PackPointer descriptor, 48, VarPtr(rhs(0)): PackPointer descriptor, 56, VarPtr(zeroCorrection(0))
    PackLong report, 0, 56: PackLong report, 4, 1
    DualRepairAssert original, solveId, candidateId, "dual_repair_before_lsmr"
    RPX_TestHook "dual_repair_before_lsmr"
    DualRepairAssert original, solveId, candidateId, "dual_repair_start_lsmr"
    started = RPX_DiagClock()
    rc = NEqualities(VarPtr(descriptor(0)), 200, 0.000000000000001, 0.000000000000001, 1000000000000#, VarPtr(correction(0)), original.equalityCount, VarPtr(report(0)))
    seconds = RPX_DiagClock() - started
    DualRepairAssert original, solveId, candidateId, "dual_repair_after_lsmr"
    If rc <> 0 Then AccuracyNativeReason = "ORIGINAL_DUAL_LINEAR_REPAIR_API_REJECTED": Exit Function
    updated = sourceDual
    For i = 0 To original.equalityCount - 1
        If Not RPX_SchurFinite(correction(i)) Then Exit Function
        updated(i) = sourceDual(i) + correction(i)
        If Not RPX_SchurFinite(updated(i)) Then Exit Function
    Next i
    unchangedCone = DualConeBitsUnchanged(sourceDual, updated, original.equalityCount, original.m)
    If Not unchangedCone Then err.Raise 5, "RPX_AccuracyNative", "DUAL_REPAIR_CHANGED_CONE_BITS"
    accepted = RPX_AccuracyDualAudit(original, updated)
    If Not RPX_SchurFinite(seconds) Or seconds < 0# Or seconds > 10# Then accepted = False
    DualRepairAssert original, solveId, candidateId, "dual_repair_after_original_audit"
    CaptureDualRepairStage original, updated, "after", accepted
    RPX_DiagEvent "original_dual_linear_repair", "accepted=" & accepted & ";Fs=" & original.factor & ";q=8;before_stat=" & beforeStat & ";before_cone=" & beforeCone & ";before_objective=" & beforeObjective & ";after_stat=" & AccuracyNativeDualStat & ";after_cone=" & AccuracyNativeDualCone & ";after_objective=" & AccuracyNativeDualObjective & ";seconds=" & seconds & ";stop=" & ReadLong(report, 8) & ";iterations=" & ReadLong(report, 12) & ";max_iterations=200;max_seconds=10;attempts=1;cone_bits_unchanged=" & unchangedCone & ";physical_bound=0;auxiliary_dual_used=0"
    If accepted Then candidate = updated: TryOriginalDualLinearRepair = True
    Exit Function
Failed:
    number = err.number: message = err.description: errorSource = err.source
    If number = 453 Then AccuracyNativeReason = "ACCURACY_DLL_EQUALITY_ENTRY_UNAVAILABLE": Exit Function
    err.Raise number, errorSource, message
End Function
'''


def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def norm(text):return text.replace('\r\n','\n').replace('\r','\n')
def changes():
    return [
        ('    Dim original As RPX_AccuracyProgram, dp As RPX_AccuracyProgram, candidate() As Double, helperDual() As Double, seconds As Double, accepted As Boolean\n',
         '    Dim original As RPX_AccuracyProgram, dp As RPX_AccuracyProgram, candidate() As Double, helperDual() As Double, seconds As Double, accepted As Boolean\n'
         '    Dim useDualRepair As Boolean, finiteHelperCandidate As Boolean, repairSolveId As Long, repairCandidateId As Long\n'),
        ('    RPX_AccuracyDualProblem original, dp\n',
         '    If mode = "upper" Then useDualRepair = DualLinearRepairEnabled()\n'
         '    repairSolveId = P4Solve: repairCandidateId = RPX_HSDCandidateId\n'
         '    RPX_AccuracyDualProblem original, dp\n'
         '    If useDualRepair Then SetUpperDualRegularization dp\n'),
        ('    If Not RPX_AccuracySolve(dp, seconds, candidate, helperDual) Then Exit Function\n',
         '    If useDualRepair Then DualRepairAssert original, repairSolveId, repairCandidateId, "original_dual_helper_before"\n'
         '    finiteHelperCandidate = RPX_AccuracySolve(dp, seconds, candidate, helperDual)\n'
         '    If useDualRepair Then DualRepairAssert original, repairSolveId, repairCandidateId, "original_dual_helper_after"\n'
         '    If Not finiteHelperCandidate Then Exit Function\n'),
        ('    accepted = RPX_AccuracyDualAudit(original, candidate)\n',
         '    accepted = RPX_AccuracyDualAudit(original, candidate)\n'
         '    If Not accepted And mode = "upper" Then accepted = TryOriginalDualLinearRepair(original, candidate)\n'),
        ('    Dim recenterSolveId As Long, recenterCandidateId As Long\n',
         '    Dim recenterSolveId As Long, recenterCandidateId As Long, useDualRepairDirect As Boolean\n'),
        ('    If mode = "upper" Then useUpperRecenter = UpperRecenterEnabled()\n',
         '    If mode = "upper" Then useUpperRecenter = UpperRecenterEnabled()\n'
         '    If mode = "upper" Then useDualRepairDirect = DualLinearRepairEnabled()\n'),
        ('    If RPX_AccuracyDualAudit(original, originalDual) Then kind = "original_dual": accepted = True\n',
         '    If RPX_AccuracyDualAudit(original, originalDual) Then kind = "original_dual": accepted = True\n'
         '    If Not accepted And useDualRepairDirect Then\n'
         '        If TryOriginalDualLinearRepair(original, originalDual) Then kind = "original_dual": accepted = True\n'
         '    End If\n'),
        ('    If recenterInvoked And number <> 0 Then err.Raise number, "RPX_AccuracyNative", message\n',
         '    If recenterInvoked And number <> 0 Then err.Raise number, "RPX_AccuracyNative", message\n'
         '    If useDualRepairDirect And number <> 0 Then err.Raise number, "RPX_AccuracyNative", message\n'),
    ]
def expected():
    text=norm(BASE.read_text(encoding='utf8'))
    for before,after in changes():
        if text.count(before)!=1:raise ValueError('Expected unique anchor: '+before)
        text=text.replace(before,after,1)
    return text+HELPERS


def main():
    p=argparse.ArgumentParser();p.add_argument('--initialize',action='store_true');args=p.parse_args()
    text=expected();src=ROOT/'src/vba/RPX_AccuracyNative.bas';imp=ROOT/'import/RPX_AccuracyNative.bas'
    if args.initialize:
        for path,enc in [(src,'utf8'),(imp,'cp932')]:
            path.parent.mkdir(parents=True,exist_ok=True)
            if not path.exists():path.write_bytes(text.replace('\n','\r\n').encode(enc))
    if norm(src.read_text(encoding='utf8'))!=text:raise ValueError('Controlled patch mismatch')
    if src.read_text(encoding='utf8')!=imp.read_text(encoding='cp932'):raise ValueError('Source/import mismatch')
    files=['build_overlay.py','README_JA.md','SETTINGS_CONTRACT.json','tests/test_contract.py','tests/native_lsmr_probe.py','tests/vba_rhs_probe.py','tests/MICROTEST_JA.md']
    result=dict(schema='rpfem.original_dual.free_y_repair.v1',status='ISOLATED_OPT_IN_PROTOTYPE; EXCEL_RUNTIME_PENDING',
                production_artifacts_modified=False,dll_modified=False,optimizer_calls=0,
                baseline=BASE.relative_to(REPO).as_posix(),baseline_sha256=sha(BASE),
                source='src/vba/RPX_AccuracyNative.bas',source_sha256=sha(src),imported='import/RPX_AccuracyNative.bas',import_sha256=sha(imp),
                profile=json.loads((ROOT/'SETTINGS_CONTRACT.json').read_text(encoding='utf8')),
                files={path:sha(ROOT/path) for path in files})
    (ROOT/'SOURCE_MANIFEST.json').write_text(json.dumps(result,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    print(json.dumps(dict(source_sha256=sha(src),manifest_sha256=sha(ROOT/'SOURCE_MANIFEST.json'))))
if __name__=='__main__':main()
