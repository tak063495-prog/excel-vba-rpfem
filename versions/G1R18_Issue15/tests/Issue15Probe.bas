Attribute VB_Name = "Issue15Probe"
Option Explicit
' Issue #15 probe: prior converged seepage result must not be accepted for coupling after a failed re-run.

Private Function CouplingState() As String
    rn = 3: re = 1
    On Error GoTo Failed
    RPX_WaterApply
    CouplingState = "ACCEPTED"
    Exit Function
Failed:
    If InStr(Err.Description, "連成：浸透") > 0 Then
        CouplingState = "REJECTED:" & Err.Description
    Else
        CouplingState = "ACCEPTED(downstream:" & Err.Description & ")"
    End If
End Function

Private Sub PrepareSheets()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("浸透入力")
    On Error GoTo 0
    If ws Is Nothing Then RPX_SeepSetupSheet: Set ws = ThisWorkbook.Worksheets("浸透入力")
    ws.Range("Z1").Value2 = "RPX_SEEP_LAYOUT_V2"
    ws.Range("D4").Value2 = 1
End Sub

Private Sub WriteConverged()
    Dim ignored As String
    ignored = RPX_SeepTestRelax(0.5)
    If Not SConverged Then Err.Raise 5, , "probe: converged reference failed"
    RPX_SeepWriteResults
End Sub

Private Function ResultFlags() As String
    With ThisWorkbook.Worksheets("浸透結果")
        ResultFlags = "K2=" & CStr(.Range("K2").Value2) & ",K9=" & CStr(.Range("K9").Value2) & ",L2=" & CStr(.Range("L2").Value2)
    End With
End Function

' kind: "validation" | "unconverged" | "success"; warm: call coupling once before re-run to fill the cache
Public Function Issue15Case(ByVal kind As String, ByVal warm As Boolean) As String
    Dim s As String, ignored As String
    PrepareSheets
    WriteConverged
    s = "initial=" & CouplingState()
    If Not warm Then
        ' cold: drop cache by writing a fresh stamp via a second reference result
        Application.Wait Now + TimeSerial(0, 0, 1)
        WriteConverged
    End If
    RPX_SeepSilent = True: RPX_SeepLastMsg = ""
    Select Case kind
        Case "validation"
            ThisWorkbook.Worksheets("浸透入力").Range("A7:Z206").ClearContents
            ThisWorkbook.Worksheets("浸透入力").Range("D4").Value2 = 1
            RPX_SeepRun
        Case "unconverged"
            ignored = RPX_SeepTestRelax(0.00000001): RPX_SeepWriteResults
        Case "success"
            Application.Wait Now + TimeSerial(0, 0, 1)
            WriteConverged
    End Select
    s = s & ";msg=" & Replace(RPX_SeepLastMsg, vbLf, " ") & ";" & ResultFlags() & ";after=" & CouplingState()
    RPX_SeepSilent = False
    Issue15Case = s
End Function
