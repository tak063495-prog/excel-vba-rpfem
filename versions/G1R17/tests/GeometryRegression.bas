Attribute VB_Name = "GeometryRegression"
Option Explicit
Public Function GeometryExport(ByVal prefix As String) As String
    Dim f As Integer, e As Long, j As Long, i As Long, v As Double, number As Long
    On Error GoTo Failed
    XCancel = False: RPX_InputEnd: RPX_ReadInputs: RPX_GenerateMesh
    MeshExport prefix & "_mesh.bin"
    f = FreeFile: Open prefix & "_physical.bin" For Binary Access Write As #f
    Put #f, , re: Put #f, , RNEdge
    For e = 0 To re - 1
        Put #f, , RRigidId(e): Put #f, , RArea(e)
        For j = 0 To 1: Put #f, , RBody0(e, j): Put #f, , RBody1(e, j): Next j
    Next e
    Close #f
    ' Exact boundary semantics and material interface node pairs.
    Dim s As Object
    Set s = CreateObject("Scripting.FileSystemObject").CreateTextFile(prefix & "_edges.txt", False, True)
    For i = 0 To RNEdge - 1
        With REdges(i)
            If .element(1) < 0 Then
                s.WriteLine .a & vbTab & .b & vbTab & .kind & vbTab & .rigidId & vbTab & RPX_ExactDouble(.load0(0)) & vbTab & RPX_ExactDouble(.load0(1)) & vbTab & RPX_ExactDouble(.load1(0)) & vbTab & RPX_ExactDouble(.load1(1))
            ElseIf RMatId(.element(0)) <> RMatId(.element(1)) Or RRigidId(.element(0)) <> RRigidId(.element(1)) Then
                s.WriteLine .a & vbTab & .b & vbTab & "interface" & vbTab & RMatId(.element(0)) & vbTab & RMatId(.element(1)) & vbTab & RRigidId(.element(0)) & vbTab & RRigidId(.element(1))
            End If
        End With
    Next i
    s.Close
    ' Reassignment must not add another fan or mix raw/reference loads.
    RPX_AssignRegions
    MeshExport prefix & "_repeat_mesh.bin"
    GeometryExport = "PASS;nodes=" & rn & ";elements=" & re & ";" & RPX_NativeInfo()
    RPX_InputEnd: Exit Function
Failed:
    number = err.number: GeometryExport = "ERROR " & number & ":" & err.description
    On Error Resume Next: Close #f: If Not s Is Nothing Then s.Close
    RPX_InputEnd
End Function
Public Function ExportCompiledPoint(ByVal side As String, ByVal fs As Double, ByVal prefix As String) As String
    On Error GoTo Failed
    XCancel = False: RPX_InputEnd: RPX_AffineEnd: RPX_ReadInputs: RPX_GenerateMesh
    RFsActive = True: RPX_TotalReferenceLoads: RStrengthFactor = fs: RProgramKind = ""
    If side = "lower" Then RPX_AssembleLower Else RPX_AssembleUpper
    RPX_ExportAssembled prefix & "_assembled.bin": MeshExport prefix & "_mesh.bin"
    ExportCompiledPoint = "PASS;side=" & side & ";variables=" & XN & ";equations=" & XM & ";cones=" & XB
    RFsActive = False: RPX_InputEnd: Exit Function
Failed:
    ExportCompiledPoint = "ERROR " & err.number & ":" & err.description
    RFsActive = False: RPX_InputEnd
End Function
