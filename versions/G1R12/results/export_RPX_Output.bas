Option Explicit

' G1R12/#1: restore before mesh tables/plots, not after writing the upper field.
Public Sub RPX_PrepareUpperOutput(ByVal adaptive As Boolean)
    RPX_AssertInputs "prepare_upper_output"
    If adaptive Then
        RPX_RestoreWitnessState RBestUpper
        RUpperField = RBestUpper.field
    End If
    RPX_AssertUpperOutput adaptive
End Sub

Public Sub RPX_AssertUpperOutput(ByVal adaptive As Boolean)
    Dim i As Long
    RPX_AssertInputs "assert_upper_output"
    If adaptive Then
        If Not RBestUpper.complete Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_OWNER_INCOMPLETE"
        If RBestUpper.inputEpoch <> RPX_InputEpoch Or RBestUpper.policy <> RPX_AnalysisPolicy() Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_INPUT_MISMATCH"
        If RBestUpper.originalModel <> RPX_ExactModel() Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_MODEL_MISMATCH"
        If RBestUpper.q <> RSubdivisions Or RBestUpper.factor <> RStrengthFactor Or RBestUpper.stressScale <> RStressScale Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_CONTEXT_MISMATCH"
    End If
    If LBound(RUpperField) <> 0 Or UBound(RUpperField) <> 12 * re + 3 * RR - 1 Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_FIELD_SHAPE"
    If adaptive Then
        If LBound(RBestUpper.field) <> 0 Or UBound(RBestUpper.field) <> UBound(RUpperField) Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_OWNER_SHAPE"
        For i = 0 To UBound(RUpperField)
            If RUpperField(i) <> RBestUpper.field(i) Then err.Raise 5, "RPX_Output", "UPPER_OUTPUT_FIELD_MISMATCH"
        Next i
    End If
End Sub

Private Sub RPX_WriteFieldsCore(ByVal adaptive As Boolean)
    Dim data() As Variant, labels() As Variant, e As Long, k As Long, j As Long, offset As Long, ws As Worksheet
    RPX_AssertUpperOutput adaptive
    RPX_InternalWrite = True
    Set ws = RPX_Sheet("要素データ")
    ReDim labels(1 To 1, 1 To 12)
    For k = 0 To 5: labels(1, 2 * k + 1) = "上界 Vx" & k + 1: labels(1, 2 * k + 2) = "上界 Vy" & k + 1: Next k
    ws.Range("H1:S1").Value2 = labels
    ReDim data(1 To re, 1 To 12)
    For e = 0 To re - 1
        For k = 0 To 11: data(e + 1, k + 1) = RUpperField(12 * e + k): Next k
    Next e
    ws.Range("H2").Resize(re, 12).Value2 = data
    ws.Range("U1").Value2 = "速度は規準化。順序：頂点1,2,3、辺12,23,31の中点。"
    If RPX_CertCurrent("lower") And RPX_CertLower.ownerKind = "IMMUTABLE_PAYLOAD" Then
        Call RPX_CertWriteLower
        GoTo upperRestore
    End If
    If adaptive Then
        RPX_RestoreWitnessState RBestLower
        RLowerField = RBestLower.field
    End If
    Set ws = RPX_Sheet("節点データ")
    ws.Range("F1:H1").Value2 = Array("下界節点ID", "X [m]", "Y [m]")
    ReDim data(1 To rn, 1 To 3)
    For e = 0 To rn - 1
        data(e + 1, 1) = e + 1: data(e + 1, 2) = RXY(e, 0): data(e + 1, 3) = RXY(e, 1)
    Next e
    ws.Range("F2").Resize(rn, 3).Value2 = data
    Set ws = RPX_Sheet("要素データ")
    ws.Range("AA1:AF1").Value2 = Array("下界要素ID", "節点1", "節点2", "節点3", "材料内部番号", "剛体内部番号")
    ReDim data(1 To re, 1 To 6)
    For e = 0 To re - 1
        data(e + 1, 1) = e + 1
        For k = 0 To 2: data(e + 1, k + 2) = RTri(e, k) + 1: Next k
        data(e + 1, 5) = RMatId(e) + 1: data(e + 1, 6) = RRigidId(e) + 1
    Next e
    ws.Range("AA2").Resize(re, 6).Value2 = data
    ReDim labels(1 To 1, 1 To 18): ReDim data(1 To re, 1 To 18)
    For k = 0 To 5
        labels(1, 3 * k + 1) = "Sxx" & k + 1 & " [kPa]": labels(1, 3 * k + 2) = "Syy" & k + 1 & " [kPa]": labels(1, 3 * k + 3) = "Txy" & k + 1 & " [kPa]"
    Next k
    If RPX_C0SplitMesh Then
        ws.Range("AA1").Value2 = "上界要素ID"
        ws.Range("AH2:AY4000").ClearContents
        ws.Range("BA1").Value2 = "下界応力は粗いメッシュの場なので、上界メッシュには書き出さない。下界Fsは結果シート。"
    Else
        For e = 0 To re - 1
            If RRigidId(e) < 0 Then
                For k = 0 To 17: data(e + 1, k + 1) = RLowerField(offset + k) * RStressScale: Next k
                offset = offset + 18
            End If
        Next e
        ws.Range("AH2").Resize(re, 18).Value2 = data
        ws.Range("BA1").Value2 = "下界応力：引張正、Bernstein係数。Fs解析では下界Fs時の場。"
    End If
    ws.Range("AH1").Resize(1, 18).Value2 = labels
upperRestore:
    If adaptive Then Call RPX_PrepareUpperOutput(True)
    RPX_InternalWrite = False
End Sub

Public Sub RPX_WriteFields(ByVal adaptive As Boolean)
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    RPX_DiagBegin dpOutput
    Call RPX_WriteFieldsCore(adaptive)
    RPX_DiagEnd dpOutput: Exit Sub
Failed:
    errorNumber = err.number: errorText = err.description
    RPX_DiagEnd dpOutput
    err.Raise errorNumber, , errorText
End Sub