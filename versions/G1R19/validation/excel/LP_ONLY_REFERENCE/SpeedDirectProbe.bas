Attribute VB_Name = "SpeedDirectProbe"
Option Explicit
Private Sub Setting(ByVal key As String, ByVal value As Variant)
    Dim ws As Worksheet, r As Long
    Set ws = RPX_Sheet("設定")
    For r = 4 To 1000
        If CStr(ws.Cells(r, 4).Value2) = key Then ws.Cells(r, 2).Value2 = value: Exit Sub
    Next r
    For r = 4 To 1000
        If Len(CStr(ws.Cells(r, 4).Value2)) = 0 And Len(CStr(ws.Cells(r, 2).Value2)) = 0 Then ws.Cells(r, 4).Value2 = key: ws.Cells(r, 2).Value2 = value: Exit Sub
    Next r
    err.Raise 5, , "SETTING_MISSING " & key
End Sub
Public Sub SpeedSetting(ByVal key As String, ByVal value As Variant)
    Setting key, value
End Sub
Public Function SpeedDirectCase(ByVal mode As String, ByVal factor As Double, ByVal direct As Long, ByVal inject As Boolean, ByVal directory As String, Optional ByVal keepInputs As Boolean = False) As Variant
    Dim value As Double, started As Double, elapsed As Double, field() As Double
    On Error GoTo Failed
    If Not keepInputs Then RPX_Setup
    Setting "TARGET", IIf(keepInputs, 400, 16): Setting "QUADRATURE", 8: Setting "RAW_LOWER_AUDIT", 1
    Setting "C0_AUTO", 0
    If Not keepInputs Then Setting "ANALYSIS_POLICY", "ASSOCIATED"
    Setting "ACCURACY_DLL_RECOVERY", 1: Setting "DIRECT_DLL", direct
    Setting "DIRECT_DLL_SECONDS", 5: Setting "HSD_SOLVE_SECONDS", 30
    Setting "TEST_INJECTION", 1: Setting "ACCURACY_DLL_CAPTURE", 1
    If Not keepInputs Then RPX_Sheet("入力").Range("B4").Value2 = 10
    RPX_Sheet("浸透入力").Range("D4").Value2 = 0
    RPX_InputBegin: RPX_ReadInputs: RPX_GenerateMesh: RPX_TotalReferenceLoads
    RLowerFs = 0.123: RUpperFs = 4.567: RBestLower.value = 0.123: RBestUpper.value = 4.567
    XLogPath = directory & "\" & mode & "_" & direct & "_" & inject & ".tsv"
    RPX_DiagStart "direct_fixed_original_lp": RFsActive = True
    AccuracyDirectLastAccepted = False: AccuracyDirectLastKind = ""
    If inject Then RPX_TestArm "direct_dll_begin", 5, "FORCED_NATIVE_FAILURE"
    started = Timer: value = RPX_Bound(mode, factor): elapsed = Timer - started
    If elapsed < 0# Then elapsed = elapsed + 86400#
    If RLowerFs <> 0.123 Or RUpperFs <> 4.567 Or RBestLower.value <> 0.123 Or RBestUpper.value <> 4.567 Then err.Raise 5, , "SAVED_BOUNDS_MUTATED"
    If inject And (AccuracyDirectLastAccepted Or XStatus <> "OPTIMAL") Then err.Raise 5, , "NATIVE_FAILURE_DID_NOT_FALL_BACK"
    SpeedDirectCase = Array("PASS", mode, factor, direct, inject, XStatus, value, elapsed, AccuracyDirectLastAccepted, AccuracyDirectLastKind, XN, XM, RSubdivisions)
    RFsActive = False: RPX_DiagFinish "PASS", "strict_original_lp_or_legacy_fallback"
    RPX_InputEnd: Exit Function
Failed:
    SpeedDirectCase = Array("FAIL", err.number, err.description, mode, factor, direct, XStatus, AccuracyNativeReason)
    RFsActive = False: RPX_InputEnd
End Function
Public Function SpeedPresetProbe() As Variant
    Dim before As Variant, materials As Variant
    On Error GoTo Failed
    RPX_Setup
    before = RPX_Sheet("要素定義").UsedRange.Value2: materials = RPX_Sheet("材料データ").UsedRange.Value2
    RPX_Speed2pctPreset
    SpeedPresetProbe = Array("PASS", RPX_Setting("GAP", 0), RPX_Setting("DIRECT_DLL", 0), RPX_Setting("QUADRATURE", 0), RPX_Setting("FS_TOL", 0), RPX_Setting("ADAPT_CERT_QUADRATURE", 0), RPX_Setting("ADAPT_PRECISION_FROM_START", -1))
    Exit Function
Failed:
    SpeedPresetProbe = Array("FAIL", err.number, err.description)
End Function
