Attribute VB_Name = "UpperRecoveryArrayProbe"
Option Explicit
' TEST ONLY: append this bridge inside a copy of RPX_AccuracyNative, removing
' Attribute/Option lines. It calls the private pure-array work helper, not a solver.
Public Function RPX_TestUpperWorkArrays() As String
    Dim pr As RPX_AccuracyProgram, candidate() As Double, ratio As Double, ok As Boolean
    Dim originalCandidate() As Double
    On Error GoTo Failed
    pr.n = 2: pr.m = 3: pr.equalityCount = 2: pr.mode = "upper": pr.q = 8: pr.phase = "reference_total"
    ReDim pr.rhs(0 To 2): pr.rhs(0) = 1#
    ReDim pr.rowPtr(0 To 3): pr.rowPtr(1) = 2: pr.rowPtr(2) = 4: pr.rowPtr(3) = 5
    ReDim pr.colId(0 To 4): pr.colId(1) = 1: pr.colId(3) = 1
    ReDim pr.values(0 To 4): pr.values(0) = 0.5: pr.values(1) = 0.5
    pr.values(2) = 1#: pr.values(3) = -1#: pr.values(4) = -1#
    ReDim candidate(0 To 1): candidate(0) = 2#: candidate(1) = 2#
    ok = NormalizeUpperWork(pr, candidate, ratio)
    If Not ok Or ratio <> 2# Or candidate(0) <> 1# Or candidate(1) <> 1# Then err.Raise 5, , "POSITIVE_WORK_IDENTITY"
    candidate(0) = 3#: candidate(1) = 1#
    ok = NormalizeUpperWork(pr, candidate, ratio)
    If Not ok Or candidate(0) = candidate(1) Then err.Raise 5, , "NORMALIZATION_MUST_NOT_REPAIR_EQUALITY"
    candidate(0) = 2#: candidate(1) = 2#: originalCandidate = candidate
    pr.rhs(1) = 0.1
    If NormalizeUpperWork(pr, candidate, ratio) Then err.Raise 5, , "TWO_NONZERO_EQUALITIES"
    If candidate(0) <> originalCandidate(0) Or candidate(1) <> originalCandidate(1) Then err.Raise 5, , "REJECT_MUTATED_CANDIDATE"
    pr.rhs(1) = 0#: pr.rhs(2) = 1E-300
    If NormalizeUpperWork(pr, candidate, ratio) Then err.Raise 5, , "NONZERO_CONE_RHS"
    pr.rhs(2) = 0#: pr.rhs(0) = 0#
    If NormalizeUpperWork(pr, candidate, ratio) Then err.Raise 5, , "ZERO_WORK_RHS"
    pr.rhs(0) = 1#: candidate(0) = -1#: candidate(1) = -1#
    If NormalizeUpperWork(pr, candidate, ratio) Then err.Raise 5, , "NEGATIVE_WORK"
    candidate(0) = 0#: candidate(1) = 0#
    If NormalizeUpperWork(pr, candidate, ratio) Then err.Raise 5, , "ZERO_WORK"
    candidate(0) = 2#: candidate(1) = 2#: pr.mode = "lower"
    If NormalizeUpperWork(pr, candidate, ratio) Then err.Raise 5, , "LOWER_SCOPE"
    pr.mode = "upper": pr.q = 4
    If NormalizeUpperWork(pr, candidate, ratio) Then err.Raise 5, , "Q4_SCOPE"
    pr.q = 8: pr.phase = "fixed_load"
    If NormalizeUpperWork(pr, candidate, ratio) Then err.Raise 5, , "PHASE_SCOPE"
    RPX_TestUpperWorkArrays = "PASS 10 pure-array identities/rejections; solver=0; certificate=0": Exit Function
Failed:
    RPX_TestUpperWorkArrays = "FAIL " & err.number & " " & err.description
End Function
