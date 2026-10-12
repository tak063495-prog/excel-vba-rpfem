Attribute VB_Name = "ContractTests"
Option Explicit
Private Sub Need(ByVal value As Boolean, ByVal text As String)
    If Not value Then err.Raise 513, , text
End Sub
Private Sub Make(ByRef ap() As Long, ByRef ai() As Long, ByRef av() As Double)
    ReDim ap(0 To 2): ReDim ai(0 To 2): ReDim av(0 To 2)
    ap(0) = 0: ap(1) = 1: ap(2) = 3: ai(0) = 0: ai(1) = 0: ai(2) = 1
    av(0) = 3#: av(1) = 1#: av(2) = 2#
End Sub
Public Function TestLUContract() As String
    Dim ap() As Long, ai() As Long, av() As Double, rhs() As Double, answer() As Double
    Dim ok As Boolean, eno As Long, text As String, mode As Variant
    On Error GoTo Failed
    For Each mode In Array("vba", "auto")
        RPX_NativeSetMode CStr(mode): XCancel = False: Make ap, ai, av
        ok = RPX_HLUFactor(ap, ai, av, 1E+9, 0#, 42)
        Need ok, "default optional LU arguments rejected"
        ReDim rhs(0 To 1): rhs(0) = 1#: rhs(1) = 2#
        RPX_HLUSolve rhs, answer, 42: Need Abs(answer(0)) < 1E-15 And Abs(answer(1) - 1#) < 1E-15, "LU default solution"
        On Error Resume Next
        RPX_HLUSolve rhs, answer, 43
        eno = err.number: text = err.description: err.Clear: On Error GoTo Failed
        Need eno = 5 And text = "HSD_FACTOR_CONTEXT_MISMATCH", "LU epoch guard"
        On Error Resume Next
        ok = RPX_HLUFactor(ap, ai, av, 1#, 0#, 42)
        eno = err.number: text = err.description: err.Clear: On Error GoTo Failed
        Need eno = 5 And text = "HSD_STABLE_UNAVAILABLE_MEMORY", "LU memory guard"
        On Error Resume Next
        ok = RPX_HLUFactor(ap, ai, av, 1E+9, 0#, 42, False, False, True, 1.1)
        eno = err.number: text = err.description: err.Clear: On Error GoTo Failed
        Need eno = 5 And text = "HSD_PIVOT_THRESHOLD_INVALID", "LU threshold guard"
        av(0) = 1E+291
        On Error Resume Next
        ok = RPX_HLUFactor(ap, ai, av, 1E+9, 0#, 42)
        eno = err.number: text = err.description: err.Clear: On Error GoTo Failed
        Need eno = 5 And text = "HSD_NONFINITE_MATRIX", "LU matrix finite predicate"
    Next mode
    RPX_NativeSetMode "auto": XCancel = True: Make ap, ai, av
    On Error Resume Next
    ok = RPX_HLUFactor(ap, ai, av, 1E+9, 0#, 42)
    eno = err.number: text = err.description: err.Clear: On Error GoTo Failed
    XCancel = False: Need eno = 18 And text = "ANALYSIS_CANCELLED", "LU cancellation"
    RPX_HLURelease
    TestLUContract = "PASS default optional arguments; shared factors; epoch; memory; threshold; finite predicate; cancellation"
    Exit Function
Failed:
    TestLUContract = "FAIL " & err.number & ":" & err.description
    XCancel = False: RPX_HLURelease
End Function

Public Function TestLUModeSwitch() As String
    Dim ap() As Long, ai() As Long, av() As Double, rhs() As Double, answer() As Double
    Dim eno As Long, text As String, mode As Variant, ok As Boolean
    On Error GoTo Failed
    For Each mode In Array("vba", "auto")
        RPX_NativeSetMode "auto": XCancel = False: Make ap, ai, av
        ok = RPX_HLUFactor(ap, ai, av, 1E+9, 0#, 77)
        Need ok, "LU mode-switch factor"
        ReDim rhs(0 To 1): rhs(0) = 1#: rhs(1) = 2#
        RPX_NativeSetMode CStr(mode)
        On Error Resume Next
        RPX_HLUSolve rhs, answer, 77
        eno = err.number: text = err.description: err.Clear: On Error GoTo Failed
        Need eno = 5 And text = "HSD_FACTOR_CONTEXT_MISMATCH", "mode-switch must invalidate the factor; got " & eno & ":" & text
        Need HLUFactorEpoch = 0, "factor epoch must be reset"
        ok = RPX_HLUFactor(ap, ai, av, 1E+9, 0#, 78): Need ok, "factor after mode switch"
        RPX_HLUSolve rhs, answer, 78
        Need Abs(answer(0)) < 1E-15 And Abs(answer(1) - 1#) < 1E-15, "fresh solution after mode switch"
    Next mode
    RPX_HLURelease
    TestLUModeSwitch = "PASS auto-to-vba and auto-to-auto invalidate stale factor; fresh solve"
    Exit Function
Failed:
    TestLUModeSwitch = "FAIL " & err.number & ":" & err.description
    RPX_HLURelease
End Function
