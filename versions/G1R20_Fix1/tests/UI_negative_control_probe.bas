Option Explicit
Public Function PublishedPresetFailure(ByVal choice As Long) As Variant
 Dim ws As Worksheet, caught As Long, leak As Boolean, meshAfter As Boolean, resultAfter As Boolean, msg As String
 Dim row As Long, key As String
 Set ws=ThisWorkbook.Worksheets("設定")
 RPX_InternalWrite=False: RPX_MeshCurrent=True: RPX_ResultCurrent=True
 If choice=3 Then
  For row=4 To 1000
   key=CStr(ws.Cells(row,4).Value2)
   Select Case key
    Case "ADAPT_BALANCED_LOWER_TARGET", "ADAPT_UPPER_FLOW_PASSES", "ACCURACY_DLL_DUAL_REGULARIZATION", "ACCURACY_DLL_DUAL_LINEAR_REPAIR", "ACCURACY_DLL_UPPER_LP_SECONDS", "ACCURACY_DLL_UPPER_RECENTER", "FS_TERMINAL_SCHEDULE"
     ws.Range(ws.Cells(row,1),ws.Cells(row,4)).ClearContents
   End Select
  Next row
 Else
  ws.Protect Password:="review-test-only", UserInterfaceOnly:=False
 End If
 On Error Resume Next
 Select Case choice
  Case 0: Call RPX_AccuracyPreset
  Case 1: Call RPX_Speed2pctPreset
  Case Else: Call RPX_G20ImprovementPreset
 End Select
 caught=Err.Number: msg=Err.Description: Err.Clear
 On Error GoTo 0
 leak=RPX_InternalWrite
 If choice<>3 Then ws.Unprotect Password:="review-test-only"
 Call RPX_Invalidate
 meshAfter=RPX_MeshCurrent: resultAfter=RPX_ResultCurrent
 PublishedPresetFailure=Array(choice,caught,msg,leak,meshAfter,resultAfter)
 RPX_InternalWrite=False: Call RPX_Invalidate
End Function
