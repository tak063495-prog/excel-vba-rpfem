Attribute VB_Name = "RPX_P4"
Option Explicit
Public Const RPX_SETTING_LAST_ROW As Long = 120

' P4 generations are advanced even when diagnostics are disabled.
Public Type RPX_DirectionInfo
    kind As Long
    attempted As Long
    returned As Long
    Matrix As Long
    solve As Long
    state As Long
    Pattern As Long
    rhs As Long
    numericValid As Boolean
    corrections As Long
    schurCorrections As Long
    checks As Long
    residual As Double
    errors(0 To 3) As Double
End Type
Public P4Trace As Long, P4Consistent As Boolean, P4Cache As Boolean, P4ReuseU As Boolean
Public P4Solve As Long, P4State As Long, P4Matrix As Long, P4Pattern As Long, P4Rhs As Long
Public P4Generic As Long, P4Current As RPX_DirectionInfo
Public P4Calls As Long, P4Replays As Long, P4Returned As Long, P4Committed As Long
Public P4Factors As Long, P4Prepares As Long, P4Hits As Long, P4Reduced As Long
Public P4Corrections As Long, P4Checks As Long, P4Attempts As Long
Public P4LocalRhs As Double, P4LocalFactor As Double
Private phaseTime(0 To 11) As Double

Private Sub SettingInvalid(ByVal key As String)
    err.Raise 5, "RPX_P4", "INVALID_SETTING " & key
End Sub
Private Function SettingNumber(ByVal key As String, ByVal value As Variant) As Double
    On Error GoTo Invalid
    If IsError(value) Or IsEmpty(value) Or VarType(value) = vbBoolean Then GoTo Invalid
    If Len(Trim$(CStr(value))) = 0 Or Not IsNumeric(value) Then GoTo Invalid
    SettingNumber = CDbl(value)
    If Not RPX_SchurFinite(SettingNumber) Then GoTo Invalid
    Exit Function
Invalid:
    err.Raise 5, "RPX_P4", "INVALID_SETTING " & key
End Function
Private Function SettingInteger(ByVal key As String, ByVal value As Variant, ByVal minimum As Double, ByVal maximum As Double) As Long
    Dim number As Double
    number = SettingNumber(key, value)
    If number <> Fix(number) Or number < minimum Or number > maximum Then Call SettingInvalid(key)
    SettingInteger = CLng(number)
End Function
Private Function SettingSwitch(ByVal key As String, ByVal value As Variant) As Long
    Dim number As Double, text As String
    If VarType(value) = vbBoolean Then
        SettingSwitch = Abs(CLng(value)): Exit Function
    End If
    If Not IsError(value) Then
        text = UCase$(Trim$(CStr(value)))
        If text = "TRUE" Then SettingSwitch = 1: Exit Function
        If text = "FALSE" Then Exit Function
        If key = "C0_AUTO" Then
            If text = "ON" Then SettingSwitch = 1: Exit Function
            If text = "OFF" Then Exit Function
        End If
    End If
    number = SettingNumber(key, value)
    If number <> 0# And number <> 1# Then Call SettingInvalid(key)
    SettingSwitch = CLng(number)
End Function
Private Function SettingChoice(ByVal key As String, ByVal value As Variant, ByVal allowed As String) As String
    Dim item As Variant, text As String
    If IsError(value) Or VarType(value) = vbBoolean Then Call SettingInvalid(key)
    text = Trim$(CStr(value))
    For Each item In Split(allowed, "|")
        If StrComp(text, CStr(item), vbTextCompare) = 0 Then SettingChoice = CStr(item): Exit Function
    Next item
    Call SettingInvalid(key)
End Function
Public Function RPX_ValidateSettingValue(ByVal key As String, ByVal value As Variant) As Variant
    Dim number As Double, countValue As Long
    key = UCase$(Trim$(key))
    If IsError(value) Then Call SettingInvalid(key)
    RPX_ValidateSettingValue = value
    Select Case key
        Case "ADAPT", "UPPER_PILOT", "UPPER_PILOT_RIGID", "RAW_LOWER_AUDIT", "LOWER_MODEL_CACHE", "SCHUR_ALLOW_RIGID", "KKT_CAPTURE", "TEST_INJECTION", "DIAGNOSTIC_MODE", "HSD_CAPTURE", "C0_AUTO", "C0_DIRECT_UPPER", "C0_UPPER_ADAPT"
            RPX_ValidateSettingValue = SettingSwitch(key, value)
        Case "DIRECT_DLL", "ACCURACY_DLL_RECOVERY", "ADAPT_PRECISION_FROM_START", "ADAPT_THIN_DLL_EQUILIBRATE", "ACCURACY_DLL_EQUILIBRATE", "ACCURACY_DLL_DUAL_EQUILIBRATE", "ACCURACY_DLL_CONE_SCALING", "ACCURACY_DLL_EQUALITY_RECOVERY", "LOWER_FIXED_RECOVERY", "ADAPT_C0_WEDGE_AUTO"
            RPX_ValidateSettingValue = SettingSwitch(key, value)
        Case "HSD_PROJECTIVE_NORMALIZATION", "HSD_ORIGINAL_OPERATOR", "HSD_LINEAR_CAPTURE", "ACCURACY_DLL_CAPTURE", "ORIGINAL_FARKAS_CAPTURE", "ACCURACY_DLL_DUAL_LINEAR_REPAIR", "ACCURACY_DLL_UPPER_RECENTER", "FS_TERMINAL_SCHEDULE", "SCHUR_STRUCTURE_CACHE"
            RPX_ValidateSettingValue = SettingSwitch(key, value)
        Case "ACCURACY_DLL_UPPER_EQUILIBRATE"
            RPX_ValidateSettingValue = SettingSwitch(key, value)
        Case "TARGET"
            RPX_ValidateSettingValue = SettingInteger(key, value, 16, 50000)
        Case "QUADRATURE"
            RPX_ValidateSettingValue = SettingInteger(key, value, 1, 16)
        Case "CYCLES", "ADAPT_ACCURACY_ROUNDS"
            RPX_ValidateSettingValue = SettingInteger(key, value, 0, 20)
        Case "ADAPT_LOWER_FAN_ROUNDS"
            RPX_ValidateSettingValue = SettingInteger(key, value, 0, 4)
        Case "ADAPT_UPPER_FLOW_PASSES"
            RPX_ValidateSettingValue = SettingInteger(key, value, 0, 5)
        Case "ADAPT_ACCURACY_CAP", "ADAPT_LOWER_TARGET", "ADAPT_BALANCED_LOWER_TARGET"
            countValue = SettingInteger(key, value, 0, 50000)
            If countValue <> 0 And countValue < 16 Then Call SettingInvalid(key)
            RPX_ValidateSettingValue = countValue
        Case "ADAPT_REFINE_CAP"
            RPX_ValidateSettingValue = SettingInteger(key, value, 16, 50000)
        Case "ADAPT_CERT_QUADRATURE"
            countValue = SettingInteger(key, value, 0, 8)
            If countValue <> 0 And countValue <> 2 And countValue <> 4 And countValue <> 8 Then Call SettingInvalid(key)
            RPX_ValidateSettingValue = countValue
        Case "SCHUR_TRACE_LEVEL", "HSD_LINEAR_TRACE"
            RPX_ValidateSettingValue = SettingInteger(key, value, 0, 2)
        Case "SCHUR_GENERIC_START_ITERS", "DIRECT_FINAL_AFTER_PILOT"
            RPX_ValidateSettingValue = SettingInteger(key, value, 0, 0)
        Case "HSD_LINEAR_CAPTURE_MAX"
            RPX_ValidateSettingValue = SettingInteger(key, value, 0, 2147483647#)
        Case "SCHUR_MAX_BYTES", "HSD_STABLE_MAX_BYTES"
            number = SettingNumber(key, value)
            If number <> Fix(number) Or number < 0# Then Call SettingInvalid(key)
            If key = "SCHUR_MAX_BYTES" And number > 2147483647# Then Call SettingInvalid(key)
            If key = "HSD_STABLE_MAX_BYTES" And number = 0# Then Call SettingInvalid(key)
            RPX_ValidateSettingValue = number
        Case "GAP", "LOWER_SEED_WIDTH", "FS_TOL", "REFINE", "SEISMIC_KH", "ADAPT_C0_WEDGE_DEPTH", "C0_GUIDE_SPACING", "ACCURACY_DLL_ALPHA", "ACCURACY_DLL_DUAL_REGULARIZATION"
            If key = "SEISMIC_KH" And Len(Trim$(CStr(value))) = 0 Then
                RPX_ValidateSettingValue = 0#: Exit Function
            End If
            number = SettingNumber(key, value)
            Select Case key
                Case "GAP": If number <= 0# Or number >= 2# Then Call SettingInvalid(key)
                Case "LOWER_SEED_WIDTH": If number <= 0# Or number >= 1# Then Call SettingInvalid(key)
                Case "FS_TOL": If number < 0.000001 Or number > 0.01 Then Call SettingInvalid(key)
                Case "REFINE", "SEISMIC_KH": If number < 0# Or number > 1# Then Call SettingInvalid(key)
                Case "ADAPT_C0_WEDGE_DEPTH", "ACCURACY_DLL_DUAL_REGULARIZATION": If number < 0# Then Call SettingInvalid(key)
                Case Else: If number <= 0# Then Call SettingInvalid(key)
            End Select
            RPX_ValidateSettingValue = number
        Case "HSD_STAGE_SECONDS"
            number = SettingNumber(key, value)
            If number < 0# Or number > 60# Then Call SettingInvalid(key)
            RPX_ValidateSettingValue = number
        Case "DIRECT_DLL_SECONDS", "ACCURACY_DLL_SECONDS", "ACCURACY_DLL_DUAL_SECONDS", "LOWER_FIXED_RECOVERY_SECONDS", "ACCURACY_DLL_UPPER_LP_SECONDS", "HSD_SOLVE_SECONDS"
            number = SettingNumber(key, value)
            If number <= 0# Then Call SettingInvalid(key)
            If key <> "HSD_SOLVE_SECONDS" And number > 600# Then Call SettingInvalid(key)
            RPX_ValidateSettingValue = number
        Case "MODE": RPX_ValidateSettingValue = SettingChoice(key, value, "Fs|極限支持力")
        Case "GRAVITY": RPX_ValidateSettingValue = SettingChoice(key, value, "固定|比例")
        Case "ANALYSIS_POLICY": RPX_ValidateSettingValue = SettingChoice(key, value, "LEGACY_DAVIS_CLIP|ASSOCIATED|DAVIS_EQUIVALENT")
        Case "LOWER_FS_SEED": RPX_ValidateSettingValue = SettingChoice(key, value, "LEGACY|UPPER_TRIAL")
        Case "LOWER_KKT_BACKEND": RPX_ValidateSettingValue = SettingChoice(key, value, "GENERIC|COMPARE|AUTO")
        Case "SCHUR_REPLAY_POLICY": RPX_ValidateSettingValue = SettingChoice(key, value, "LEGACY|CONSISTENT")
        Case "SCHUR_RECOVERY_MODE": RPX_ValidateSettingValue = SettingChoice(key, value, "RESOLVE|REUSE_U")
        Case "SCHUR_REFINEMENT": RPX_ValidateSettingValue = SettingChoice(key, value, "FIXED3")
        Case "HSD_LINEAR_BACKEND": RPX_ValidateSettingValue = SettingChoice(key, value, "LDL_CHECKED|PIVOT_ONLY|AUTO_PIVOT")
        Case "ADAPT_MESH_POLICY": RPX_ValidateSettingValue = SettingChoice(key, value, "LEGACY|ROBUST_REFINE|LEGACY_THEN_REFINE")
        Case "FS_BRACKET_POLICY": RPX_ValidateSettingValue = SettingChoice(key, value, "LEGACY|AUDITED_UPPER_LIMIT")
        Case "C0_MESH_ROUTE": RPX_ValidateSettingValue = SettingChoice(key, value, "BAND|CHORD")
        Case "SEISMIC_DIR": RPX_ValidateSettingValue = SettingChoice(key, Replace(Replace(CStr(value), "（", "("), "）", ")"), "右(+x)|左(-x)|右|左|+x|-x|x|+X|-X|X|")
    End Select
End Function
Public Sub RPX_ValidateSettingsExtent()
    Dim ws As Worksheet, last As Long
    Set ws = RPX_Sheet("設定")
    ' End(xlUp) includes hidden keys and formulas, even when their value is "".
    ' Description-only cells in A:C are outside the keyed model and remain notes.
    last = ws.cells(ws.rows.count, 4).End(xlUp).row
    If last > RPX_SETTING_LAST_ROW Then err.Raise 5, "RPX_P4", "SETTING_ROW_OUT_OF_RANGE row=" & last
End Sub
Public Sub RPX_ValidateSettings()
    Dim ws As Worksheet, seen As Object, row As Long, raw As Variant, key As String, ignored As Variant
    Call RPX_ValidateSettingsExtent
    Set ws = RPX_Sheet("設定"): Set seen = CreateObject("Scripting.Dictionary")
    seen.CompareMode = vbTextCompare
    For row = 4 To RPX_SETTING_LAST_ROW
        raw = RPX_InputCell(ws, row, 4)
        If IsError(raw) Then err.Raise 5, "RPX_P4", "INVALID_SETTING_KEY row=" & row
        If Not IsEmpty(raw) Then
            If VarType(raw) <> vbString Then err.Raise 5, "RPX_P4", "INVALID_SETTING_KEY row=" & row
            key = UCase$(Trim$(CStr(raw)))
            If Len(key) > 0 Then
                If seen.Exists(key) Then err.Raise 5, "RPX_P4", "DUPLICATE_SETTING " & key
                seen.Add key, row
                ignored = RPX_ValidateSettingValue(key, RPX_InputCell(ws, row, 2))
            End If
        End If
    Next row

End Sub


Private Function Choice(ByVal key As String, ByVal fallback As String, ByVal allowed As String) As String
    Choice = UCase$(Trim$(CStr(RPX_Setting(key, fallback))))
    If InStr("|" & allowed & "|", "|" & Choice & "|") = 0 Or Len(Choice) = 0 Then err.Raise 5, "RPX_P4", "INVALID_SETTING " & key
End Function
Public Sub RPX_P4Configure()
    Dim value As String
    P4Trace = CLng(Choice("SCHUR_TRACE_LEVEL", "0", "0|1|2"))
    P4Consistent = (Choice("SCHUR_REPLAY_POLICY", "LEGACY", "LEGACY|CONSISTENT") = "CONSISTENT")
    P4Cache = (Choice("SCHUR_STRUCTURE_CACHE", "0", "0|1") = "1")
    P4ReuseU = (Choice("SCHUR_RECOVERY_MODE", "RESOLVE", "RESOLVE|REUSE_U") = "REUSE_U")
    value = Choice("SCHUR_REFINEMENT", "FIXED3", "FIXED3")
    value = Choice("SCHUR_GENERIC_START_ITERS", "0", "0")
    value = Choice("DIRECT_FINAL_AFTER_PILOT", "0", "0")
End Sub
Public Sub RPX_P4Begin()
    Dim i As Long
    RPX_P4Configure
    ' Reserve more than a complete solve before any Long counter can overflow.
    If P4Solve > 2000000000 Or P4Matrix > 2000000000 Or P4Rhs > 2000000000 Or P4Pattern > 2000000000 Or P4State > 2000000000 Then
        RPX_SchurRelease
        P4Solve = 0: P4Matrix = 0: P4Rhs = 0: P4Pattern = 0: P4State = 0
    End If
    P4Solve = P4Solve + 1: P4Generic = -1
    P4Calls = 0: P4Replays = 0: P4Returned = 0: P4Committed = 0
    P4Factors = 0: P4Prepares = 0: P4Hits = 0: P4Reduced = 0
    P4Corrections = 0: P4Checks = 0: P4Attempts = 0
    P4LocalRhs = 0#: P4LocalFactor = 0#
    For i = 0 To 11: phaseTime(i) = 0#: Next i
    RPX_DiagEvent "p4_engine", "engine=20260927_p4_3;trace=" & P4Trace & ";consistent=" & P4Consistent & ";structure_cache=" & P4Cache & ";reuse_u=" & P4ReuseU & ";refinement=FIXED3;generic_start_iters=0;direct_final=0"
End Sub
Public Sub RPX_P4NewMatrix()
    P4Matrix = P4Matrix + 1: P4Generic = -1
End Sub
Public Sub RPX_P4StartDirection(ByVal kind As Long)
    Dim blank As RPX_DirectionInfo
    P4Current = blank: P4Rhs = P4Rhs + 1: P4Calls = P4Calls + 1
    P4Current.solve = P4Solve: P4Current.state = P4State: P4Current.Pattern = P4Pattern
    P4Current.kind = kind: P4Current.Matrix = P4Matrix: P4Current.rhs = P4Rhs
End Sub
Public Sub RPX_P4Time(ByVal phase As Long, ByVal started As Double)
    phaseTime(phase) = phaseTime(phase) + RPX_DiagClock() - started
End Sub
Public Sub RPX_P4Direction(ByVal backend As Long, ByVal residual As Double)
    P4Current.returned = backend: P4Current.numericValid = True: P4Current.residual = residual
    If P4Trace = 2 Then RPX_DiagEvent "p4_residual_history", "rhs=" & P4Current.rhs & ";e0=" & P4Current.errors(0) & ";e1=" & P4Current.errors(1) & ";e2=" & P4Current.errors(2) & ";e3=" & P4Current.errors(3)
    If backend = 2 Then P4Returned = P4Returned + 1
    P4Checks = P4Checks + P4Current.checks
    If P4Trace > 0 Then RPX_DiagEvent "p4_direction_return", "solve=" & P4Solve & ";state=" & P4State & ";matrix=" & P4Matrix & ";pattern=" & P4Pattern & ";rhs=" & P4Current.rhs & ";kind=" & P4Current.kind & ";attempted=" & P4Current.attempted & ";returned=" & backend & ";corrections=" & P4Current.corrections & ";schur_corrections=" & P4Current.schurCorrections & ";checks=" & P4Current.checks & ";residual=" & residual
End Sub
Public Sub RPX_P4Commit(ByRef affine As RPX_DirectionInfo, ByRef Corrector As RPX_DirectionInfo)
    If affine.kind <> 1 Or Corrector.kind <> 2 Or affine.returned < 1 Or affine.returned > 2 Then err.Raise 5, "RPX_P4", "DIRECTION_PAIR_KIND"
    If affine.solve <> P4Solve Or Corrector.solve <> P4Solve Or affine.state <> P4State Or Corrector.state <> P4State Then err.Raise 5, "RPX_P4", "DIRECTION_PAIR_STATE"
    If Not affine.numericValid Or Not Corrector.numericValid Or affine.Matrix <> P4Matrix Or Corrector.Matrix <> P4Matrix Or affine.returned <> Corrector.returned Then err.Raise 5, "RPX_P4", "DIRECTION_PAIR_IDENTITY"
    If affine.returned = 2 Then P4Committed = P4Committed + 2
    If P4Trace > 0 Then RPX_DiagEvent "p4_pair_commit", "matrix=" & P4Matrix & ";affine_rhs=" & affine.rhs & ";corrector_rhs=" & Corrector.rhs & ";backend=" & affine.returned
    P4State = P4State + 1
End Sub
Public Sub RPX_P4Summary()
    Dim i As Long, names As Variant
    RPX_DiagEvent "p4_summary", "calls=" & P4Calls & ";replays=" & P4Replays & ";schur_returned=" & P4Returned & ";schur_committed=" & P4Committed & ";fallback_factors=" & P4Factors & ";prepare=" & P4Prepares & ";hits=" & P4Hits & ";reduced_solves=" & P4Reduced & ";corrections=" & P4Corrections & ";checks=" & P4Checks & ";schur_attempts=" & P4Attempts & ";local_rhs_solves=" & P4LocalRhs & ";local_factor_solves=" & P4LocalFactor
    names = Split("prepare,local_factor,schur_update,numeric_only,reduced_rhs,backsolve,recover,full_residual,cache_verify,correction_inclusive,replay_inclusive,fallback_factor", ",")
    For i = 0 To 11
        RPX_DiagEvent "p4_time", "phase=" & names(i) & ";seconds=" & phaseTime(i) & ";inclusive=" & Abs(CLng(i = 9 Or i = 10))
    Next i
End Sub
