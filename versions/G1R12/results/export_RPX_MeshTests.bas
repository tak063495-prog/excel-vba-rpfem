Option Explicit
Private Function DefinitionRejected() As Boolean
    On Error GoTo rejected
    Call RPX_PrepareDefinition: Exit Function
rejected:
    DefinitionRejected = True
End Function
Public Function RPX_TestDefinitionValidation() As String
    Dim e As Long, k As Long, id As Long, a0 As Double, a1 As Double, interfaces As Long
    On Error GoTo Failed
    RPX_SlopeDefinition 64: DFaces(2).wrap = False
    If Not DefinitionRejected() Then err.Raise 5, , "WRAP_ERROR_NOT_REJECTED"
    RPX_SlopeDefinition 64: DRegions(0).points(2) = DRegions(0).points(1)
    If Not DefinitionRejected() Then err.Raise 5, , "DUPLICATE_POINT_NOT_REJECTED"
    DPCount = 6: DRCount = 2: DFCount = 0: RM = 2: RR = 0: DTarget = 64
    ReDim DPXY(0 To 5, 0 To 1): ReDim DRegions(0 To 1): ReDim RMaterials(0 To 1)
    DPXY(0, 0) = 0#: DPXY(0, 1) = 0#: DPXY(1, 0) = 0#: DPXY(1, 1) = 1#
    DPXY(2, 0) = 1#: DPXY(2, 1) = 1#: DPXY(3, 0) = 1#: DPXY(3, 1) = 0#
    DPXY(4, 0) = 2#: DPXY(4, 1) = 1#: DPXY(5, 0) = 2#: DPXY(5, 1) = 0#
    For e = 0 To 1
        DRegions(e).pointCount = 4: DRegions(e).materialId = e: DRegions(e).rigidId = -1
        ReDim DRegions(e).points(0 To 3): RMaterials(e).cohesion = e + 1#
    Next e
    For k = 0 To 3: DRegions(0).points(k) = k: Next k
    DRegions(1).points(0) = 3: DRegions(1).points(1) = 2: DRegions(1).points(2) = 4: DRegions(1).points(3) = 5
    Call RPX_GenerateMesh
    For e = 0 To re - 1
        If RMatId(e) = 0 Then a0 = a0 + RArea(e) Else a1 = a1 + RArea(e)
    Next e
    If Abs(a0 - 1#) + Abs(a1 - 1#) > 0.000001 Then err.Raise 5, , "MATERIAL_AREA_FAILED"
    For id = 0 To RNEdge - 1
        With REdges(id)
            If .element(1) >= 0 Then
                If RMatId(.element(0)) <> RMatId(.element(1)) Then
                    interfaces = interfaces + 1
                    If Abs(RXY(.a, 0) - 1#) + Abs(RXY(.b, 0) - 1#) > 0.000001 Then err.Raise 5, , "MATERIAL_INTERFACE_FAILED"
                End If
            End If
        End With
    Next id
    If interfaces = 0 Then err.Raise 5, , "MATERIAL_INTERFACE_MISSING"
    RPX_TestDefinitionValidation = "PASS invalid paths rejected, adjacent material boundary preserved": Exit Function
Failed:
    RPX_TestDefinitionValidation = "FAIL " & err.description
End Function
Public Function RPX_TestIntermediatePoint() As String
    On Error GoTo Failed
    RPX_SlopeDefinition 64
    ' VBA cannot Preserve the first dimension; copy this small definition explicitly.
    Dim points(0 To 5, 0 To 1) As Double, i As Long, j As Long
    For i = 0 To 5: For j = 0 To 1: points(i, j) = DPXY(i, j): Next j: Next i
    DPCount = 7: ReDim DPXY(0 To 6, 0 To 1)
    For i = 0 To 5: For j = 0 To 1: DPXY(i, j) = points(i, j): Next j: Next i
    DPXY(6, 0) = 9#: DPXY(6, 1) = 15#
    DFaces(0).startPoint = 6: DFaces(0).endPoint = 3: DFaces(0).kind = "load"
    Call RPX_PrepareDefinition
    If DRegions(0).pointCount <> 7 Then err.Raise 5, , "INTERMEDIATE_POINT_NOT_INSERTED"
    If Not RPX_FaceContains(0, 12#, 15#) Or Not RPX_FaceContains(0, 33#, 25#) Or RPX_FaceContains(0, 4#, 15#) Then err.Raise 5, , "FACE_PATH_FAILED"
    Call RPX_GenerateMesh
    RPX_TestIntermediatePoint = "PASS intermediate endpoint and bent path": Exit Function
Failed:
    RPX_TestIntermediatePoint = "FAIL " & err.description
End Function
Public Sub RPX_SlopeDefinition(ByVal target As Long)
    Dim i As Long, p As Variant
    p = Array(Array(0#, 0#), Array(0#, 15#), Array(18#, 15#), Array(48#, 35#), Array(66#, 35#), Array(66#, 0#))
    DPCount = 6: DRCount = 1: DFCount = 3: DTarget = target: DSpacing = 0#
    ReDim DPXY(0 To 5, 0 To 1): ReDim DRegions(0 To 0): ReDim DFaces(0 To 2)
    For i = 0 To 5: DPXY(i, 0) = p(i)(0): DPXY(i, 1) = p(i)(1): Next i
    DRegions(0).pointCount = 6: ReDim DRegions(0).points(0 To 5)
    For i = 0 To 5: DRegions(0).points(i) = i: Next i
    DRegions(0).materialId = 0: DRegions(0).rigidId = -1
    For i = 0 To 2: DFaces(i).regionId = 0: DFaces(i).kind = "fixed": DFaces(i).rigidId = -1: Next i
    DFaces(0).startPoint = 0: DFaces(0).endPoint = 1
    DFaces(1).startPoint = 4: DFaces(1).endPoint = 5
    DFaces(2).startPoint = 5: DFaces(2).endPoint = 0: DFaces(2).wrap = True
    RM = 1: RR = 0: ReDim RMaterials(0 To 0): ReDim RRigids(0 To 0)
    RMaterials(0).cohesion = 41.65: RMaterials(0).friction = 15#: RMaterials(0).gamma = 18.82
    RStressScale = 41.65: RStrengthFactor = 1#: RSubdivisions = 8: RGravityMode = "fixed"
End Sub
Public Function RPX_TestMesh(ByVal target As Long, ByVal outputPath As String) As String
    Dim f As Integer, i As Long, area As Double
    On Error GoTo Failed
    XLogPath = outputPath & ".log": RPX_SlopeDefinition target: Call RPX_GenerateMesh
    For i = 0 To re - 1: area = area + RArea(i): Next i
    If Abs(area - 1650#) > 0.0000001 Then err.Raise 5, , "MESH_AREA_MISMATCH"
    f = FreeFile: Open outputPath For Output As #f
    Print #f, rn; ","; re; ","; area
    For i = 0 To rn - 1: Print #f, RXY(i, 0); ","; RXY(i, 1): Next i
    For i = 0 To re - 1: Print #f, RTri(i, 0); ","; RTri(i, 1); ","; RTri(i, 2): Next i
    Close #f: RPX_TestMesh = "PASS nodes=" & rn & " elements=" & re: Application.StatusBar = False: Exit Function
Failed:
    RPX_TestMesh = "FAIL " & err.number & " " & err.description: Application.StatusBar = False
End Function