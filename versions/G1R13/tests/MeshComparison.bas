Attribute VB_Name = "MeshComparison"
Option Explicit
Public Function CompareMesh(ByVal path As String) As String
    Dim fs As Object, s As Object, area As Double, weight As Double, e As Long
    Call RPX_InputEnd: XCancel = False: RPX_ReadInputs: DTarget = 538: RPX_GenerateMesh
    For e = 0 To re - 1
        area = area + RArea(e)
        weight = weight - RArea(e) * (RBody0(e, 1) + RBody1(e, 1))
    Next e
    Set fs = CreateObject("Scripting.FileSystemObject"): Set s = fs.CreateTextFile(path, False, True)
    s.WriteLine RPX_ExactModel(): s.Close
    CompareMesh = rn & vbTab & re & vbTab & area & vbTab & weight & vbTab & RPX_LoadPhase
End Function
