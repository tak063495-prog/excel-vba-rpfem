Attribute VB_Name = "RPX_NativeGuard"
Option Explicit
#If Win64 Then
Private Declare PtrSafe Function NCompare Lib "rpfem_native_x64.dll" Alias "rpx_snapshot_compare" (ByVal live As LongPtr, ByVal forms As LongPtr, ByVal saved As LongPtr, ByVal oldforms As LongPtr, ByVal columns As LongPtr, ByVal n As Long, ByVal rows As Long, ByRef changed As Long) As Long
#End If
Private calls As Double
Public Function RPX_NativeSnapshot(ByRef live As Variant, ByRef forms As Variant, ByRef saved As Variant, ByRef oldforms As Variant, ByRef columns As Variant, ByVal n As Long, ByVal rows As Long, ByRef changed As Boolean) As Boolean
    If Not RPX_NativeAvailable() Then Exit Function
#If Win64 Then
    Dim code As Long, difference As Long
    code = NCompare(VarPtr(live), VarPtr(forms), VarPtr(saved), VarPtr(oldforms), VarPtr(columns), n, rows, difference)
    If code = 3 Then Exit Function
    If code = 330 Then err.Raise 9, "RPX_Guard", "NATIVE_SNAPSHOT_SHAPE_MISMATCH"
    RPX_NativeCheckCode code
    calls = calls + 1: changed = (difference <> 0): RPX_NativeSnapshot = True
#End If
End Function
Public Function RPX_NativeGuardInfo() As String
    RPX_NativeGuardInfo = ";input_compare_calls=" & calls
End Function
Public Sub RPX_NativeGuardResetCounts()
    calls = 0#
End Sub
