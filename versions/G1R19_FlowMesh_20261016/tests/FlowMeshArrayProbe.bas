Attribute VB_Name = "FlowMeshArrayProbe"
Option Explicit
' Test-only binary bridge. No model installation and no solver call.
Public Function FlowMeshFileProbe(ByVal inputPath As String, ByVal outputPath As String) As String
    Dim f As Integer, magic As String * 8, nn As Long, ne As Long, guardCount As Long, passes As Long
    Dim xy() As Double, tri() As Long, materialId() As Long, velocity() As Double, candidate() As Long
    Dim oldXY() As Double, oldTri() As Long, oldMat() As Long, oldVelocity() As Double
    Dim guards As Object, info As Object, changes As Collection, item As Object, edge As Variant, owners As Variant
    Dim i As Long, j As Long, a As Long, b As Long, number As Long, value As Double, flag As Long
    Dim accepted As Boolean, message As String, started As Double, elapsed As Double, partialPath As String
    On Error GoTo Failed
    If Len(Dir$(outputPath)) > 0 Or inputPath = outputPath Then err.Raise 5, , "FLOW_PROBE_REFUSES_OVERWRITE"
    f = FreeFile: Open inputPath For Binary Access Read As #f
    Get #f, , magic
    If magic <> "FMG1I001" Then err.Raise 5, , "FLOW_PROBE_INPUT_MAGIC"
    Get #f, , nn: Get #f, , ne: Get #f, , guardCount: Get #f, , passes
    If nn < 3 Or nn > 100000 Or ne < 1 Or ne > 220000 Or guardCount < 0 Or guardCount > 3 * ne Then err.Raise 5, , "FLOW_PROBE_INPUT_SIZE"
    ReDim xy(0 To nn - 1, 0 To 1): ReDim tri(0 To ne - 1, 0 To 2)
    ReDim materialId(0 To ne - 1): ReDim velocity(0 To 12 * ne - 1)
    For i = 0 To nn - 1: For j = 0 To 1: Get #f, , value: xy(i, j) = value: Next j: Next i
    For i = 0 To ne - 1: For j = 0 To 2: Get #f, , number: tri(i, j) = number: Next j: Next i
    For i = 0 To ne - 1: Get #f, , number: materialId(i) = number: Next i
    For i = 0 To 12 * ne - 1: Get #f, , value: velocity(i) = value: Next i
    Set guards = CreateObject("Scripting.Dictionary")
    For i = 0 To guardCount - 1
        Get #f, , a: Get #f, , b
        If a >= b Or guards.Exists(CStr(a) & ":" & CStr(b)) Then err.Raise 5, , "FLOW_PROBE_GUARD_KEY"
        guards.Add CStr(a) & ":" & CStr(b), True
    Next i
    If Seek(f) <> LOF(f) + 1 Then err.Raise 5, , "FLOW_PROBE_TRAILING_BYTES"
    Close #f: f = 0
    oldXY = xy: oldTri = tri: oldMat = materialId: oldVelocity = velocity
    started = Timer
    accepted = RPX_FlowMeshCandidate(xy, tri, materialId, velocity, guards, candidate, info, passes)
    elapsed = Timer - started: If elapsed < 0# Then elapsed = elapsed + 86400#
    If Not accepted Or CBool(info("certificate")) Or CBool(info("nested")) Or Not CBool(info("requires_new_original_q8_field_and_root")) Then err.Raise 5, , "FLOW_PROBE_NOT_GEOMETRY_ONLY"
    For i = 0 To nn - 1: For j = 0 To 1: If xy(i, j) <> oldXY(i, j) Then err.Raise 5, , "FLOW_PROBE_POINT_MUTATED"
    Next j: Next i
    For i = 0 To ne - 1
        If materialId(i) <> oldMat(i) Then err.Raise 5, , "FLOW_PROBE_MATERIAL_MUTATED"
        For j = 0 To 2: If tri(i, j) <> oldTri(i, j) Then err.Raise 5, , "FLOW_PROBE_SOURCE_TRIANGLE_MUTATED"
        Next j
    Next i
    For i = 0 To 12 * ne - 1: If velocity(i) <> oldVelocity(i) Then err.Raise 5, , "FLOW_PROBE_SOURCE_VELOCITY_MUTATED"
    Next i
    If guards.count <> guardCount Then err.Raise 5, , "FLOW_PROBE_GUARDS_MUTATED"
    Set changes = info("changes"): partialPath = outputPath & ".partial"
    If Len(Dir$(partialPath)) > 0 Then err.Raise 5, , "FLOW_PROBE_PARTIAL_EXISTS"
    f = FreeFile: Open partialPath For Binary Access Write As #f
    magic = "FMG1O001": Put #f, , magic: Put #f, , nn: Put #f, , ne
    number = changes.count: Put #f, , number: flag = 0: Put #f, , flag: Put #f, , elapsed
    For i = 0 To ne - 1: For j = 0 To 2: number = candidate(i, j): Put #f, , number: Next j: Next i
    For Each item In changes
        edge = item("old"): number = CLng(edge(0)): Put #f, , number: number = CLng(edge(1)): Put #f, , number
        edge = item("new"): number = CLng(edge(0)): Put #f, , number: number = CLng(edge(1)): Put #f, , number
        owners = item("elements"): number = CLng(owners(0)): Put #f, , number: number = CLng(owners(1)): Put #f, , number
        number = CLng(item("sweep")): Put #f, , number: value = CDbl(item("gain")): Put #f, , value
    Next item
    Close #f: f = 0: Name partialPath As outputPath
    FlowMeshFileProbe = "PASS geometry_only;flips=" & changes.count & ";passes=" & passes & ";seconds=" & elapsed & ";solver_calls=0;certificate=False"
    Exit Function
Failed:
    message = err.description
    On Error Resume Next
    If f <> 0 Then Close #f
    FlowMeshFileProbe = "FAIL " & message
End Function
