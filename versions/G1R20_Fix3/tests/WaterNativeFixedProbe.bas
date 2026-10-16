Attribute VB_Name = "WaterNativeFixedProbe"
Option Explicit
' Test-only fresh canonical D08 points. No saved geometry/field/Fs/root import.
Private Sub RequireWater(ByVal accepted As Boolean, ByVal reason As String)
    If Not accepted Then Err.Raise 5, "WaterNativeFixedProbe", reason
End Sub
Private Sub CheckWaterGeometry(ByVal faces As Long)
    Dim e As Long, k As Long, id As Long, below As Boolean, area As Double, expected As Double, key As String, waterLength As Double
    RequireWater RPX_WaterCommonApplied And RPX_WaterCommonConforming(3#), "WATER_COMMON_NOT_APPLIED"
    RequireWater DFCount = faces And RSubdivisions = 8 And RR = 0, "WATER_GEOMETRY_CONTEXT"
    For e = 0 To re - 1
        area = area + RArea(e): below = True
        For k = 0 To 2: If RXY(RTri(e, k), 1) > 3# Then below = False
        Next k
        expected = -RMaterials(RMatId(e)).gamma
        If below Then expected = expected + 9.81
        RequireWater Abs(RBody0(e, 0)) < 0.000000000001 And Abs(RBody1(e, 0)) < 0.000000000001, "WATER_HORIZONTAL_BODY"
        RequireWater RBody0(e, 1) = 0# And Abs(RBody1(e, 1) - expected) < 0.000000000001, "WATER_REFERENCE_BODY"
    Next e
    RequireWater Abs(area - RPX_UnionArea()) < 0.00000001, "ORIGINAL_REGION_AREA"
    For id = 0 To RNEdge - 1
        With REdges(id)
            If .element(1) >= 0 And RXY(.a, 1) = 3# And RXY(.b, 1) = 3# Then
                key = RPX_EdgeKey(.a, .b)
                waterLength = waterLength + Abs(RXY(.a, 0) - RXY(.b, 0))
                RequireWater RConstraints.Exists(key) And Not RBoundary.Exists(key), "WATERLINE_PHYSICAL_BOUNDARY"
                RequireWater .kind = "load" And .rigidId = -1, "WATERLINE_RESTRAINT"
                For k = 0 To 1: RequireWater .load0(k) = 0# And .load1(k) = 0#, "WATERLINE_FORCE": Next k
            End If
        End With
    Next id
    RequireWater Abs(waterLength - 20#) < 0.00000001, "WATERLINE_INTERNAL_LENGTH"
End Sub
Private Sub WriteWaterGeometry(ByVal path As String)
    Dim f As Integer, i As Long, j As Long, e As Long, value As Variant, key As Variant, ends As Variant, kind As Long, v As Double
    f = FreeFile: Open path For Binary Access Write As #f
    RPX_ExchangeMagic f, "RP25G001"
    Put #f, , rn: Put #f, , re: Put #f, , RM: Put #f, , RR: Put #f, , RSubdivisions: Put #f, , RStressScale
    For i = 0 To rn - 1: For j = 0 To 1: Put #f, , RXY(i, j): Next j: Next i
    For e = 0 To re - 1
        For j = 0 To 2: Put #f, , RTri(e, j): Next j
        Put #f, , RMatId(e): Put #f, , RRigidId(e)
    Next e
    For i = 0 To RM - 1
        Put #f, , RMaterials(i).cohesion: Put #f, , RMaterials(i).friction
        Put #f, , RMaterials(i).dilation: Put #f, , RMaterials(i).gamma
    Next i
    For e = 0 To re - 1: For j = 0 To 1: Put #f, , RBody0(e, j): Put #f, , RBody1(e, j): Next j: Next e
    i = RBoundary.Count: Put #f, , i
    For Each key In RBoundary.Keys
        ends = Split(CStr(key), ":"): i = CLng(ends(0)): j = CLng(ends(1))
        value = RBoundary(key)
        Select Case CStr(value(0))
            Case "load": kind = 0
            Case "fixed": kind = 1
            Case "roller_x": kind = 2
            Case "roller_y": kind = 3
            Case Else: Err.Raise 5, , "WATER_BOUNDARY_KIND"
        End Select
        Put #f, , i: Put #f, , j: Put #f, , kind: i = CLng(value(1)): Put #f, , i
        For j = 2 To 5: v = CDbl(value(j)): Put #f, , v: Next j
    Next key
    Close #f
End Sub
Public Function WaterNativeFixedPoint(ByVal side As String, ByVal factor As Double, ByVal prefix As String) As Variant
    Dim faces As Long, original As RPX_AccuracyProgram, field() As Double, objective As Double
    Dim measured As Double, audit As Double, raw() As Double, a As Double, b As Double, accepted As Boolean, guided As Boolean
    Dim number As Long, message As String
    On Error GoTo Failed
    RequireWater side = "lower" Or side = "upper", "SIDE"
    Call RPX_InputBegin: Call RPX_ReadInputs
    DTarget = 64: RSubdivisions = 8: faces = DFCount
    Call RPX_GenerateMesh: Call RPX_TotalReferenceLoads
    Call CheckWaterGeometry(faces): Call WriteWaterGeometry(prefix & "_geometry.bin")
    Call RPX_FsBracketReset
    XLogPath = prefix & ".trace.tsv": Call RPX_DiagStart("fix1_water_fixed_point")
    RFsActive = True: RStrengthFactor = factor: RPX_OrderMode = side
    Call RPX_CapacityAssembly(side)
    If side = "lower" Then Call RPX_AssembleLower Else Call RPX_AssembleUpper
    Call RPX_ExportAssembled(prefix & "_original_assembled.bin")
    Call RPX_ResetPreparedState: Call RPX_Compile
    RequireWater RPX_AccuracyExport(factor, side, original), "ORIGINAL_EXPORT"
    accepted = RPX_AccuracyDirectBound(factor, side)
    guided = accepted And RPX_AccuracyGuideCurrent(factor, side)
    RequireWater accepted And Not guided, "NO_NATIVE_ORIGINAL_PRIMAL"
    field = xx
    RequireWater RPX_AccuracyPrimalAudit(original, field, objective), "RAW_ORIGINAL_PRIMAL"
    RequireWater RPX_FieldGate(side, field, factor, 8, objective, measured, audit, raw), "ORIGINAL_PHYSICAL_FIELD"
    If side = "lower" Then RequireWater measured >= 1.000001, "LOWER_WITNESS" Else RequireWater measured <= 0.999999, "UPPER_WITNESS"
    Call RPX_AccuracyWriteProof(prefix & "_accepted_original_primal.bin", original, field)
    RequireWater Not RPX_ActualRootBracket(side, a, b, factor), "FIXED_POINT_CREATED_ROOT"
    WaterNativeFixedPoint = Array("PASS", side, factor, measured, objective, re, rn, AccuracyNativeEquality, AccuracyNativeCone, "root_not_claimed", "fresh_canonical_D08")
Cleanup:
    If number = 0 Then Call RPX_DiagFinish("PASS") Else Call RPX_DiagFinish("FAIL", message)
    RFsActive = False: Call RPX_InputEnd
    Exit Function
Failed:
    number = Err.Number: message = Err.Description
    WaterNativeFixedPoint = Array("FAIL", side, factor, number, message)
    Resume Cleanup
End Function
