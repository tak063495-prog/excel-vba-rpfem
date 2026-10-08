Option Explicit
Public RBearingWidth As Double, RBearingFixed As Double, RBearingReference As Double
Public RLowerLoad As Double, RUpperLoad As Double
Private measureValid As Boolean, measureModel As String, measureSelection As String

' Selected surface loads only. Positive result is downward force per unit depth.
' Mean pressure is vertical resultant / horizontal projected contact width.
Private Sub BearingMeasureCore(ByVal selection As String)
    Dim selected As Object, parts As Variant, part As Variant, f As Long, id As Long
    Dim edge As Long, a As Long, b As Long, x As Double, y As Double
    Dim dx As Double, dy As Double, length As Double, covered As Boolean, rigidSelection As String
    Set selected = RPX_Row()
    If RPX_FaceIDs Is Nothing Then err.Raise 5, , "面定義を読み込んでください。"
    If Len(Trim$(selection)) = 0 Then err.Raise 5, , "支持圧集計の面IDを指定してください。"
    parts = Split(Replace(selection, "、", ","), ",")
    For Each part In parts
        If Not IsNumeric(Trim$(CStr(part))) Then err.Raise 5, , "支持圧集計の面IDが不正です。"
        x = CDbl(part)
        If x < 1# Or x > 1000000# Or x <> Fix(x) Then err.Raise 5, , "支持圧集計の面IDが不正です。"
        id = CLng(x)
        If Not RPX_FaceIDs.Exists(id) Then err.Raise 5, , "支持圧集計の面IDが未定義です。"
        selected(CLng(RPX_FaceIDs(id))) = True
    Next part
    RBearingWidth = 0#: RBearingFixed = 0#: RBearingReference = 0#
    For edge = 0 To RNEdge - 1
        If REdges(edge).element(1) < 0 Then
            a = REdges(edge).a: b = REdges(edge).b
            dx = RXY(b, 0) - RXY(a, 0): dy = RXY(b, 1) - RXY(a, 1)
            length = Sqr(dx * dx + dy * dy)
            x = (RXY(a, 0) + RXY(b, 0)) / 2#: y = (RXY(a, 1) + RXY(b, 1)) / 2#
            covered = False
            For f = 0 To DFCount - 1
                If selected.Exists(f) Then
                    If RPX_FaceContains(f, x, y) Then
                        covered = True
                        RBearingFixed = RBearingFixed - length * DFaces(f).load0(1)
                        RBearingReference = RBearingReference - length * DFaces(f).load1(1)
                    End If
                End If
            Next f
            If covered Then RBearingWidth = RBearingWidth + Abs(dx)
        End If
    Next edge
    rigidSelection = Trim$(CStr(RPX_Setting("BEARING_RIGIDS", "")))
    If Len(rigidSelection) > 0 Then
        If RPX_RigidIDs Is Nothing Then err.Raise 5, , "剛体定義を読み込んでください。"
        Set selected = RPX_Row()
        For Each part In Split(Replace(rigidSelection, "、", ","), ",")
            If Not IsNumeric(Trim$(CStr(part))) Then err.Raise 5, , "支持圧集計の剛体IDが不正です。"
            x = CDbl(part)
            If x < 1# Or x > 1000000# Or x <> Fix(x) Then err.Raise 5, , "支持圧集計の剛体IDが不正です。"
            id = CLng(x)
            If Not RPX_RigidIDs.Exists(id) Then err.Raise 5, , "支持圧集計の剛体IDが未定義です。"
            If Not selected.Exists(id) Then
                selected(id) = True: f = RPX_RigidIDs(id)
                RBearingFixed = RBearingFixed - RRigids(f).load0(1)
                RBearingReference = RBearingReference - RRigids(f).load1(1)
            End If
        Next part
    End If
    If RBearingWidth <= DTolerance Then err.Raise 5, , "支持圧集計面の水平投影幅がゼロです。"
    If RBearingReference <= 0# Then err.Raise 5, , "支持圧集計面には下向きの比例荷重を指定してください。"
End Sub

Public Function RPX_BearingPressure(ByVal multiplier As Double) As Double
    RPX_BearingAssertCurrent
    If RBearingWidth <= 0# Then err.Raise 5, , "支持圧集計が未設定です。"
    RPX_BearingPressure = (RBearingFixed + multiplier * RBearingReference) / RBearingWidth
End Function

Public Sub RPX_LoadPair()
    Dim i As Long
    RUpperLoad = RPX_Bound("upper", 1#)
    ReDim RUpperField(0 To 12 * re + 3 * RR - 1)
    For i = 0 To UBound(RUpperField): RUpperField(i) = xx(i): Next i
    RLowerLoad = RPX_Bound("lower", 1#)
    ReDim RLowerField(0 To UBound(xx))
    For i = 0 To UBound(RLowerField): RLowerField(i) = xx(i): Next i
    If RLowerLoad > RUpperLoad Then err.Raise 5, , "LOAD_BOUNDS_REVERSED"
    If RLowerLoad < 0# Then err.Raise 5, , "固定荷重だけで限界に達する可能性があります。"
End Sub

Public Function RPX_TestBearingMeasure() As String
    ' Bent face with an intermediate definition point; duplicate selection must not double width.
    RPX_SlopeDefinition 64
    DFaces(0).startPoint = 1: DFaces(0).endPoint = 4: DFaces(0).kind = "load"
    DFaces(0).load0(1) = -2#: DFaces(0).load1(1) = -3#
    Call RPX_PrepareDefinition
    Call RPX_GenerateMesh
    Set RPX_FaceIDs = RPX_Row(): RPX_FaceIDs(7&) = 0&
    RPX_BearingMeasure "7,7"
    If Abs(RBearingWidth - 66#) > 0.000001 Then err.Raise 5, , "BEARING_WIDTH_FAILED"
    If Abs(RPX_BearingPressure(4#) - 14# * (36# + Sqr(1300#)) / 66#) > 0.000001 Then err.Raise 5, , "BEARING_PRESSURE_FAILED"
    RPX_TestBearingMeasure = "PASS width=" & CStr(RBearingWidth) & " pressure=" & CStr(RPX_BearingPressure(4#))
End Function

Public Sub RPX_BearingInvalidate()
    measureValid = False: measureModel = "": measureSelection = ""
    RBearingWidth = 0#: RBearingFixed = 0#: RBearingReference = 0#
End Sub
Private Sub ValidateIDs(ByVal selection As String, ByVal mapping As Object, ByVal required As Boolean)
    Dim part As Variant, value As Double
    If Len(Trim$(selection)) = 0 Then
        If required Then err.Raise 5, "RPX_Bearing", "BEARING_FACE_SELECTION_REQUIRED"
        Exit Sub
    End If
    If mapping Is Nothing Then err.Raise 5, "RPX_Bearing", "BEARING_DEFINITION_REQUIRED"
    For Each part In Split(Replace(selection, "、", ","), ",")
        If Not IsNumeric(Trim$(CStr(part))) Then err.Raise 5, "RPX_Bearing", "INVALID_BEARING_ID"
        value = CDbl(part)
        If value < 1# Or value > 1000000# Or value <> Fix(value) Then err.Raise 5, "RPX_Bearing", "INVALID_BEARING_ID"
        If Not mapping.Exists(CLng(value)) Then err.Raise 5, "RPX_Bearing", "UNDEFINED_BEARING_ID"
    Next part
End Sub
Public Sub RPX_BearingValidateSelection()
    ValidateIDs CStr(RPX_Setting("BEARING_FACES", "")), RPX_FaceIDs, True
    ValidateIDs CStr(RPX_Setting("BEARING_RIGIDS", "")), RPX_RigidIDs, False
    RPX_DiagEvent "bearing_selection_valid", "geometry_required=False"
End Sub
Public Sub RPX_BearingMeasure(ByVal selection As String)
    Dim number As Long, message As String, source As String, line As Long
    On Error GoTo Failed
    RPX_BearingInvalidate
    RPX_AssertInputs "bearing_measure"
    If RPX_LoadPhase <> "raw" Then err.Raise 5, "RPX_Bearing", "BEARING_LOAD_PHASE_MISMATCH"
    BearingMeasureCore selection
    measureModel = RPX_ExactModel()
    measureSelection = selection & "|" & CStr(RPX_Setting("BEARING_RIGIDS", ""))
    measureValid = True
    RPX_DiagEvent "bearing_measure", "load_phase=" & RPX_LoadPhase & ";nodes=" & rn & ";elements=" & re & ";width=" & RBearingWidth & ";fixed=" & RBearingFixed & ";reference=" & RBearingReference
    Exit Sub
Failed:
    number = err.number: message = err.description: source = err.source: line = Erl
    RPX_BearingInvalidate
    RPX_ErrorEvidence number, source, message, line, "bearing_measure"
    err.Raise number, source, message
End Sub
Public Sub RPX_BearingAssertCurrent()
    RPX_AssertInputs "bearing_result"
    If Not measureValid Or RPX_LoadPhase <> "raw" Then err.Raise 5, "RPX_Bearing", "BEARING_CONTEXT_INVALID"
    If measureModel <> RPX_ExactModel() Then err.Raise 5, "RPX_Bearing", "BEARING_MODEL_MISMATCH"
    If RPX_Busy Then
        If measureSelection <> CStr(RPX_Setting("BEARING_FACES", "")) & "|" & CStr(RPX_Setting("BEARING_RIGIDS", "")) Then err.Raise 5, "RPX_Bearing", "BEARING_SELECTION_MISMATCH"
    End If
End Sub