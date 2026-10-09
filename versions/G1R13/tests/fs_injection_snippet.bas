' Test-only procedures appended to RPX_Fs in an isolated controls workbook.
Public Function FsRepairInjectedTrial(ByVal number As Long, ByVal message As String) As String
    Dim accepted As Boolean, value As Double, beforeX() As Double, i As Long
    Dim source As String, actual As Long, actualMessage As String
    On Error GoTo Failed
    XCancel = False: RProgramKind = "sentinel_model": XStatus = "sentinel_status"
    ReDim xx(0 To 1): xx(0) = 11#: xx(1) = 22#: beforeX = xx: value = 123#
    RPX_TestArm "fs_value", number, message
    accepted = TryFsValue(1.2, "upper", 0.999999, value)
    If accepted Or value <> 123# Or xx(0) <> 11# Or xx(1) <> 22# Then err.Raise 5, "FsRepairInjectedTrial", "UNKNOWN_ADOPTED_OR_FIELD_CHANGED"
    FsRepairInjectedTrial = "UNKNOWN" & vbTab & XStatus & vbTab & RProgramKind & vbTab & RPX_TestHits & vbTab & value
    Exit Function
Failed:
    actual = err.number: source = err.source: actualMessage = err.description
    FsRepairInjectedTrial = "ERROR" & vbTab & actual & vbTab & source & vbTab & actualMessage & vbTab & RPX_TestHits
    XCancel = False: RPX_TestDisarm
End Function
