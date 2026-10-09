Attribute VB_Name = "LoadStepControls"
Option Explicit
Private Sub SetFan(ByVal pressureLeft As Double, ByVal pressureRight As Double, ByVal innerX As Double, Optional ByVal angle As Double = 0#, Optional ByVal fixedLeft As Boolean = False)
    Dim i As Long, j As Long, x As Double, y As Double, co As Double, si As Double
    XCancel = False: RProgramKind = "": RPX_LoadPhase = "raw"
    rn = 4: re = 2: RM = 1: RR = 0: RStressScale = 1#: RGravityMode = "fixed"
    ReDim RXY(0 To 31, 0 To 1): ReDim RTri(0 To 31, 0 To 2): ReDim RMaterials(0 To 0)
    ReDim RMatId(0 To 1): ReDim RRigidId(0 To 1): ReDim RBody0(0 To 1, 0 To 1): ReDim RBody1(0 To 1, 0 To 1)
    RRigidId(0) = -1: RRigidId(1) = -1: RMaterials(0).cohesion = 10#: RMaterials(0).friction = 25#: RMaterials(0).dilation = 25#: RMaterials(0).gamma = 20#
    RXY(1, 0) = -1#: RXY(2, 0) = 1#: RXY(3, 0) = innerX: RXY(3, 1) = -1#
    co = Cos(angle): si = Sin(angle)
    For i = 0 To 3
        x = RXY(i, 0): y = RXY(i, 1): RXY(i, 0) = co * x - si * y: RXY(i, 1) = si * x + co * y
    Next i
    RPX_SetPositive 0, 0, 1, 3: RPX_SetPositive 1, 0, 3, 2
    Set RConstraints = RPX_Row(): Set RBoundary = RPX_Row()
    RBoundary(RPX_EdgeKey(0, 1)) = Array(IIf(fixedLeft, "fixed", "load"), -1, si * pressureLeft, -co * pressureLeft, 0#, 0#)
    RBoundary(RPX_EdgeKey(0, 2)) = Array("load", -1, si * pressureRight, -co * pressureRight, 0#, 0#)
    RBoundary(RPX_EdgeKey(1, 3)) = Array("fixed", -1, 0#, 0#, 0#, 0#)
    RBoundary(RPX_EdgeKey(2, 3)) = Array("fixed", -1, 0#, 0#, 0#, 0#)
    RPX_Topology
End Sub
Public Function LoadStepTest(ByVal kind As String) As String
    Dim changed As Boolean, expected As Boolean, area As Double, e As Long, original As String, savedXY() As Double, i As Long, j As Long
    On Error GoTo Failed
    Select Case kind
        Case "step": SetFan 0#, 20#, 0.3: expected = True
        Case "reversed": SetFan 20#, 0#, 0.3: expected = True
        Case "equal": SetFan 20#, 20#, 0.3
        Case "zero": SetFan 0#, 0#, 0.3
        Case "vertical": SetFan 0#, 20#, 0#
        Case "rotated": SetFan 0#, 20#, 0.3, 0.7: expected = True
        Case "fixed": SetFan 0#, 20#, 0.3, 0#, True
        Case "capacity"
            SetFan 0#, 20#, 0.3: savedXY = RXY: ReDim RXY(0 To 6, 0 To 1)
            For i = 0 To 3
                For j = 0 To 1: RXY(i, j) = savedXY(i, j): Next j
            Next i
        Case "phase": SetFan 0#, 20#, 0.3: RPX_LoadPhase = "reference_total"
        Case "cancel": SetFan 0#, 20#, 0.3: XCancel = True
        Case "subscript": SetFan 0#, 20#, 0.3: RTri(1, 2) = 999
        Case Else: err.Raise 5, "LoadStepControls", "UNKNOWN_CONTROL"
    End Select
    original = RPX_ExactModel()
    changed = RPX_RepairLoadSteps()
    If changed <> expected Then err.Raise 5, "LoadStepControls", "UNEXPECTED_REPAIR_DECISION"
    If expected Then
        If rn <> 6 Or re <> 6 Then err.Raise 5, "LoadStepControls", "REPAIR_COUNTS"
        For e = 0 To re - 1: area = area + RPX_Cross(RTri(e, 0), RTri(e, 1), RTri(e, 2)) / 2#: Next e
        If Abs(area - 1#) > 0.000000000001 Then err.Raise 5, "LoadStepControls", "REPAIR_AREA"
    Else
        If RPX_ExactModel() <> original Then err.Raise 5, "LoadStepControls", "NO_REPAIR_IDENTITY"
    End If
    LoadStepTest = "PASS" & vbTab & kind & vbTab & changed & vbTab & rn & vbTab & re
    Exit Function
Failed:
    LoadStepTest = "ERROR" & vbTab & kind & vbTab & err.number & vbTab & err.description: XCancel = False
End Function
