Attribute VB_Name = "SettingsInputFix4Probe"
Option Explicit
Private Function SettingRow(ByVal key As String) As Long
    Dim r As Long
    For r = 4 To 100
        If CStr(RPX_Sheet("設定").Cells(r, 4).Value2) = key Then SettingRow = r: Exit Function
    Next r
    For r = 100 To 4 Step -1
        If IsEmpty(RPX_Sheet("設定").Cells(r, 4).Value2) Then
            RPX_Sheet("設定").Cells(r, 4).Value2 = key: SettingRow = r: Exit Function
        End If
    Next r
    Err.Raise 5, , "TEST_NO_SETTING_ROW"
End Function
Public Function SettingsRoute(ByVal value As Variant) As Variant
    Dim r As Long, number As Long, message As String, route As Boolean
    On Error GoTo Failed
    XCancel = False: RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_Setup
    RPX_Sheet("材料データ").Range("B5").Value2 = 0#
    r = SettingRow("C0_AUTO"): RPX_Sheet("設定").Cells(r, 2).Value2 = value
    Call RPX_ReadInputs: route = RPX_C0AutoRequested()
    RPX_InternalWrite = False
    SettingsRoute = Array("ACCEPTED", route, VarType(value), CStr(value)): Exit Function
Failed:
    number = Err.Number: message = Err.Description: RPX_InternalWrite = False
    SettingsRoute = Array("REJECTED", number, message)
End Function
Public Function SettingsQuadrature(ByVal value As Variant) As Variant
    Dim r As Long, number As Long, message As String, met As Boolean
    On Error GoTo Failed
    XCancel = False: RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_Setup
    r = SettingRow("ADAPT_CERT_QUADRATURE"): RPX_Sheet("設定").Cells(r, 2).Value2 = value
    Call RPX_ReadInputs: Call RPX_CertClear
    RFsSearchIncomplete = False: RPX_AuditRejected = False
    met = RPX_TargetMet("Fs", 1#, 1#, True)
    RPX_InternalWrite = False
    SettingsQuadrature = Array("ACCEPTED", met, VarType(value), CStr(value)): Exit Function
Failed:
    number = Err.Number: message = Err.Description: RPX_InternalWrite = False
    SettingsQuadrature = Array("REJECTED", number, message)
End Function
Public Function SettingsDuplicate() As Variant
    Dim r As Long, number As Long, message As String
    On Error GoTo Failed
    XCancel = False: RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_Setup
    For r = 100 To 4 Step -1
        If IsEmpty(RPX_Sheet("設定").Cells(r, 4).Value2) Then Exit For
    Next r
    RPX_Sheet("設定").Cells(r, 4).Value2 = "TARGET": RPX_Sheet("設定").Cells(r, 2).Value2 = 4096#
    Call RPX_ReadInputs
    RPX_InternalWrite = False
    SettingsDuplicate = Array("ACCEPTED", DTarget, r, RPX_Setting("TARGET", -1)): Exit Function
Failed:
    number = Err.Number: message = Err.Description: RPX_InternalWrite = False
    SettingsDuplicate = Array("REJECTED", number, message)
End Function

Public Function SettingsOutOfRange() As Variant
    Dim r As Long, number As Long, message As String
    On Error GoTo Failed
    XCancel = False: RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_Setup
    For r = 4 To 120
        If CStr(RPX_Sheet("設定").Cells(r, 4).Value2) = "SEISMIC_KH" Or CStr(RPX_Sheet("設定").Cells(r, 4).Value2) = "SEISMIC_DIR" Then RPX_Sheet("設定").Cells(r, 4).ClearContents
    Next r
    RPX_Sheet("設定").Range("D110").Value2 = "SEISMIC_KH": RPX_Sheet("設定").Range("B110").Value2 = 0.2
    RPX_Sheet("設定").Range("D111").Value2 = "SEISMIC_DIR": RPX_Sheet("設定").Range("B111").Value2 = "右(+x)"
    Call RPX_SeismicSetup: Call RPX_ReadInputs
    SettingsOutOfRange = Array("ACCEPTED", RPX_SeismicKhSigned(), RPX_Sheet("設定").Range("B110").Value2, RPX_Sheet("設定").Range("D110").Value2)
    RPX_InternalWrite = False: Exit Function
Failed:
    number = Err.Number: message = Err.Description: RPX_InternalWrite = False
    SettingsOutOfRange = Array("REJECTED", number, message)
End Function

Public Function SettingsMissingSeismicKey(ByVal which As String) As Variant
    Dim rowKh As Long, rowDir As Long, r As Long, countKh As Long, countDir As Long, number As Long, message As String
    On Error GoTo Failed
    XCancel = False: RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_Setup
    rowKh = SettingRow("SEISMIC_KH"): rowDir = SettingRow("SEISMIC_DIR")
    RPX_Sheet("設定").Cells(rowKh, 2).Value2 = 0.2
    RPX_Sheet("設定").Cells(rowDir, 2).Value2 = "左(-x)"
    If which = "missing_dir" Then
        RPX_Sheet("設定").Cells(rowDir, 4).ClearContents
    Else
        RPX_Sheet("設定").Cells(rowKh, 4).ClearContents
    End If
    Call RPX_SeismicSetup
    For r = 4 To 130
        If CStr(RPX_Sheet("設定").Cells(r, 4).Value2) = "SEISMIC_KH" Then countKh = countKh + 1
        If CStr(RPX_Sheet("設定").Cells(r, 4).Value2) = "SEISMIC_DIR" Then countDir = countDir + 1
    Next r
    SettingsMissingSeismicKey = Array("ACCEPTED", RPX_SeismicKhSigned(), RPX_Setting("SEISMIC_DIR", "MISSING"), countKh, countDir)
    RPX_InternalWrite = False: Exit Function
Failed:
    number = Err.Number: message = Err.Description: RPX_InternalWrite = False
    SettingsMissingSeismicKey = Array("REJECTED", number, message)
End Function

Public Function SettingsGeneral(ByVal key As String, ByVal value As Variant) As Variant
    Dim r As Long, number As Long, message As String, accepted As Variant
    On Error GoTo Failed
    XCancel = False: RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_Setup
    r = SettingRow(key): RPX_Sheet("設定").Cells(r, 2).Value2 = value
    Call RPX_ReadInputs: accepted = RPX_Setting(key, "MISSING")
    SettingsGeneral = Array("ACCEPTED", accepted): RPX_InternalWrite = False: Exit Function
Failed:
    number = Err.Number: message = Err.Description: RPX_InternalWrite = False
    SettingsGeneral = Array("REJECTED", number, message)
End Function
Public Function SettingsUnsupportedRow() As Variant
    Dim number As Long, message As String
    On Error GoTo Failed
    XCancel = False: RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_Setup
    RPX_Sheet("設定").Range("D121").Value2 = "SEISMIC_KH": RPX_Sheet("設定").Range("B121").Value2 = 0.2
    Call RPX_ReadInputs
    SettingsUnsupportedRow = Array("ACCEPTED", RPX_SeismicKhSigned()): RPX_InternalWrite = False: Exit Function
Failed:
    number = Err.Number: message = Err.Description: RPX_InternalWrite = False
    SettingsUnsupportedRow = Array("REJECTED", number, message)
End Function
Public Function SettingsExtendedEvent(ByVal formulaInput As Boolean) As Variant
    Dim unused As Variant, ref As Worksheet, r As Long, number As Long, message As String
    On Error GoTo Failed
    unused = SettingsOutOfRange(): RPX_InternalWrite = True
    If formulaInput Then
        Set ref = RPX_Sheet("参照入力"): ref.Range("A1").Value2 = 0.2
        RPX_Sheet("設定").Range("B110").Formula = "='参照入力'!A1": RPX_Sheet("設定").Calculate
    End If
    RPX_Sheet("解析結果").Range("A1").Value2 = "SYNTHETIC_SAVED_SENTINEL"
    RPX_Sheet("解析結果").Range("B4:C4").Value2 = Array(1.11, 2.22)
    RPX_ResultCurrent = True: RPX_MeshCurrent = True: Call RPX_InputObserve
    RPX_InternalWrite = False
    If formulaInput Then
        ref.Range("A1").Value2 = 0.3: Application.CalculateFull
    Else
        RPX_Sheet("設定").Range("B110").Value2 = 0.3
    End If
    SettingsExtendedEvent = Array("ACCEPTED", RPX_ResultCurrent, RPX_Sheet("解析結果").Range("B4").Value2, RPX_Sheet("設定").Range("B110").Value2)
    Exit Function
Failed:
    number = Err.Number: message = Err.Description: RPX_InternalWrite = False
    SettingsExtendedEvent = Array("REJECTED", number, message)
End Function
Public Function SettingsGapDefaults() As Variant
    Dim value As Double, display As String, pass15 As Boolean, pass21 As Boolean, explicitValue As Double, detail As Double, speed As Double, improve As Double
    On Error GoTo Failed
    XCancel = False: RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_Setup
    value = RPX_Setting("GAP", -1): display = RPX_Sheet("設定").Range("B9").Text
    RPX_Sheet("設定").Range("D9").ClearContents: Call RPX_ReadInputs
    Call RPX_CertClear: RFsSearchIncomplete = False: RPX_AuditRejected = False
    pass15 = RPX_TargetMet("Fs", 0.9925, 1.0075, True)
    pass21 = RPX_TargetMet("Fs", 0.9895, 1.0105, True)
    RPX_Sheet("設定").Range("D9").Value2 = "GAP": RPX_Sheet("設定").Range("B9").Value2 = 0.01
    explicitValue = RPX_Setting("GAP", -1)
    Call RPX_AccuracyPreset: detail = RPX_Setting("GAP", -1)
    Call RPX_Speed2pctPreset: speed = RPX_Setting("GAP", -1)
    Call RPX_G20ImprovementPreset: improve = RPX_Setting("GAP", -1)
    RPX_InternalWrite = False
    SettingsGapDefaults = Array("ACCEPTED", value, display, pass15, pass21, explicitValue, detail, speed, improve)
    Exit Function
Failed:
    SettingsGapDefaults = Array("REJECTED", Err.Number, Err.Description): RPX_InternalWrite = False
End Function

Public Function SettingsSeismicStable(ByVal scenario As String) As Variant
    Dim ws As Worksheet, rowKh As Long, rowDir As Long, before As String, after As String, nKh As Long, nDir As Long, r As Long
    On Error GoTo Failed
    XCancel = False: RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_Setup
    Set ws = RPX_Sheet("設定")
    If scenario <> "both_missing" Then
        Call RPX_SeismicSetup
        rowKh = SettingRow("SEISMIC_KH"): rowDir = SettingRow("SEISMIC_DIR")
        ws.Cells(rowKh, 2).Formula = "=0.2"
        ws.Cells(rowDir, 2).Formula = "=""左(-x)"""
        If scenario = "missing_dir_formula" Then ws.Cells(rowDir, 4).ClearContents
        If scenario = "missing_kh_formula" Then ws.Cells(rowKh, 4).ClearContents
    End If
    before = RPX_SourceRange(ws.Range("A4:D120"))
    Call RPX_SeismicSetup
    after = RPX_SourceRange(ws.Range("A4:D120"))
    rowKh = SettingRow("SEISMIC_KH"): rowDir = SettingRow("SEISMIC_DIR")
    For r = 4 To 120
        If CStr(ws.Cells(r, 4).Value2) = "SEISMIC_KH" Then nKh = nKh + 1
        If CStr(ws.Cells(r, 4).Value2) = "SEISMIC_DIR" Then nDir = nDir + 1
    Next r
    SettingsSeismicStable = Array("ACCEPTED", RPX_SeismicKhSigned(), ws.Cells(rowKh, 2).Formula, ws.Cells(rowDir, 2).Formula, nKh, nDir, before = after)
    RPX_InternalWrite = False: Exit Function
Failed:
    SettingsSeismicStable = Array("REJECTED", Err.Number, Err.Description): RPX_InternalWrite = False
End Function

Public Function SettingsMetadata(ByVal scenario As String) As Variant
    Dim ws As Worksheet, row As Long
    On Error GoTo Failed
    XCancel = False: RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_Setup
    Set ws = RPX_Sheet("設定")
    Select Case scenario
        Case "blank_formula_tail": ws.Range("D121").Formula = "="""""
        Case "notes_tail": ws.Range("A121:C121").Value2 = Array("利用者メモ", 0.2, "説明のみ")
        Case "error_key": ws.Range("D120").Formula = "=NA()"
        Case "numeric_key": ws.Range("D120").Value2 = 123
        Case "canonical_keys"
            row = SettingRow("HSD_SOLVE_SECONDS"): ws.Cells(row, 4).Value2 = " hsd_solve_seconds "
            row = SettingRow("C0_AUTO"): ws.Cells(row, 4).Value2 = " c0_auto "
            ws.Range("D5").Value2 = " target ": ws.Range("B5").Value2 = 32
            Call RPX_EnsureUpperPilotSetting: Call RPX_C0EnsureSetting
    End Select
    Call RPX_ReadInputs
    SettingsMetadata = Array("ACCEPTED", RPX_Setting("TARGET", -1)): RPX_InternalWrite = False: Exit Function
Failed:
    SettingsMetadata = Array("REJECTED", Err.Number, Err.Description): RPX_InternalWrite = False
End Function

Public Function SettingsExcelError(ByVal key As String) As Variant
    Dim row As Long
    On Error GoTo Failed
    XCancel = False: RPX_InternalWrite = True: Call RPX_InputEnd: Call RPX_Setup
    row = SettingRow(key): RPX_Sheet("設定").Cells(row, 2).Formula = "=NA()"
    Call RPX_ReadInputs
    SettingsExcelError = Array("ACCEPTED", key): RPX_InternalWrite = False: Exit Function
Failed:
    SettingsExcelError = Array("REJECTED", Err.Number, Err.Description): RPX_InternalWrite = False
End Function
