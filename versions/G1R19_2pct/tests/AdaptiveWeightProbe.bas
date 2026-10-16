Attribute VB_Name = "AdaptiveWeightProbe"
Option Explicit
Public Function AdaptiveWeightUnit() As Variant
    Dim field() As Double, chosen() As Double, e As Long, k As Long, m As Long
    Dim a As Long, b As Long, expected As Double, zeroTotal As Double, c0Total As Double
    Dim positiveTotal As Double, originalC() As Double, originalPhi() As Double, c As Double, phi As Double
    On Error GoTo Failed
    RPX_Setup
    OwnedSetting "TARGET", 16
    OwnedSetting "C0_AUTO", 0
    RPX_InputBegin: RPX_ReadInputs: RPX_GenerateMesh
    ReDim originalC(0 To RM - 1): ReDim originalPhi(0 To RM - 1)
    For m = 0 To RM - 1
        originalC(m) = RMaterials(m).cohesion: originalPhi(m) = RMaterials(m).friction
        RMaterials(m).cohesion = 0#: RMaterials(m).friction = 0.5
    Next m
    ReDim field(0 To 12 * re - 1)
    For e = 0 To re - 1
        For k = 0 To 5
            If k < 3 Then
                a = RTri(e, k): b = a
            Else
                a = RTri(e, k - 3): b = RTri(e, (k - 2) Mod 3)
            End If
            field(12 * e + 2 * k) = (RXY(a, 0) + RXY(b, 0)) / 2#
            field(12 * e + 2 * k + 1) = (RXY(a, 1) + RXY(b, 1)) / 2#
        Next k
    Next e
    RPX_AdaptiveRefineWeight field, chosen
    For e = 0 To re - 1
        expected = RArea(e) * Sqr(2#)
        If Abs(chosen(e) - expected) > 0.000000001 * (1# + expected) Then Err.Raise 5, , "C0_KINEMATIC_WEIGHT"
        c0Total = c0Total + chosen(e)
    Next e
    If c0Total <= 0# Then Err.Raise 5, , "C0_WEIGHT_STILL_ZERO"
    For m = 0 To RM - 1
        If RMaterials(m).cohesion <> 0# Then Err.Raise 5, , "MATERIAL_CHANGED_BY_WEIGHT"
        RMaterials(m).cohesion = 5#
    Next m
    RPX_AdaptiveRefineWeight field, chosen
    For e = 0 To re - 1
        RPX_Strength e, c, phi
        expected = RArea(e) * 2# * c / Tan(phi)
        If Abs(chosen(e) - expected) > 0.000000001 * (1# + expected) Then Err.Raise 5, , "POSITIVE_C_WEIGHT_CHANGED"
        positiveTotal = positiveTotal + chosen(e)
    Next e
    Erase field: ReDim field(0 To 12 * re - 1)
    RPX_AdaptiveRefineWeight field, chosen
    For e = 0 To re - 1: zeroTotal = zeroTotal + chosen(e): Next e
    If zeroTotal <> 0# Then Err.Raise 5, , "ZERO_FIELD_CREATED_WEIGHT"
    AdaptiveWeightUnit = Array("PASS", re, c0Total, positiveTotal, zeroTotal)
CleanUp:
    On Error Resume Next
    For m = 0 To RM - 1
        RMaterials(m).cohesion = originalC(m): RMaterials(m).friction = originalPhi(m)
    Next m
    RPX_InputEnd
    Exit Function
Failed:
    AdaptiveWeightUnit = Array("FAIL", Err.Number, Err.Description)
    Resume CleanUp
End Function

Public Function AdaptiveWedgeGuardUnit() As Variant
    Dim oldC As Double, oldLoad As Double, oldX As Double, beforeElements As Long
    Dim e As Long, area As Double, originalArea As Double, beforeLower As Double, beforeUpper As Double
    On Error GoTo Failed
    RPX_InputBegin: RPX_ReadInputs: RPX_GenerateMesh
    beforeElements = re: oldC = RMaterials(0).cohesion
    oldLoad = DFaces(0).load0(0): oldX = DPXY(4, 0)
    beforeLower = RBestLower.value: beforeUpper = RBestUpper.value
    originalArea = DRegions(0).area
    If RPX_C0SlopeWedgeMesh(0#) Then Err.Raise 5, , "ZERO_DEPTH_ACCEPTED"
    RMaterials(0).cohesion = 1#
    If RPX_C0SlopeWedgeMesh(0.006) Then Err.Raise 5, , "COHESIVE_WEDGE_ACCEPTED"
    RMaterials(0).cohesion = oldC: DFaces(0).load0(0) = 1#
    If RPX_C0SlopeWedgeMesh(0.006) Then Err.Raise 5, , "LOADED_WEDGE_ACCEPTED"
    DFaces(0).load0(0) = oldLoad: DPXY(4, 0) = DPXY(3, 0)
    If RPX_C0SlopeWedgeMesh(0.006) Then Err.Raise 5, , "UNSUPPORTED_GEOMETRY_ACCEPTED"
    DPXY(4, 0) = oldX
    If re <> beforeElements Or RBestLower.value <> beforeLower Or RBestUpper.value <> beforeUpper Then Err.Raise 5, , "REJECTION_CHANGED_STATE"
    If Not RPX_C0SlopeWedgeMesh(0.001 * (DPXY(4, 1) - DPXY(3, 1))) Then Err.Raise 5, , "SUPPORTED_GENERIC_WEDGE_REJECTED"
    For e = 0 To re - 1: area = area + RArea(e): Next e
    If re <> 8 Or area <> originalArea Or RMaterials(0).cohesion <> oldC Then Err.Raise 5, , "WEDGE_ORIGINAL_MODEL_CHANGED"
    AdaptiveWedgeGuardUnit = Array("PASS", beforeElements, re, area, "material_load_geometry_rejections_atomic")
CleanUp:
    On Error Resume Next
    RMaterials(0).cohesion = oldC: DFaces(0).load0(0) = oldLoad: DPXY(4, 0) = oldX
    RPX_InputEnd
    Exit Function
Failed:
    AdaptiveWedgeGuardUnit = Array("FAIL", Err.Number, Err.Description)
    Resume CleanUp
End Function
