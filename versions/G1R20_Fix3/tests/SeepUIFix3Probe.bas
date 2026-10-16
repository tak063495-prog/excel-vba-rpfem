Attribute VB_Name = "SeepUIFix3Probe"
Option Explicit
' Disposable Excel copies only. Mode2 prescribed hydrostatics, no optimizer.
Public Sub SeepUIFix3Prepare()
    Dim ws As Worksheet
    XCancel = False: RPX_InternalWrite = False: RPX_Busy = False
    Call RPX_InputEnd
    Call RPX_Setup
    Call RPX_SeepSetSilent(True)
    Call RPX_SeepSetupSheet
    RPX_Sheet("設定").Range("B5").Value2 = 16
    Set ws = RPX_Sheet("浸透入力")
    ws.Range("B3").Value2 = ws.Range("Z6").Value2
    ws.Range("D4").Value2 = 0
    ws.Range("M7:N206").ClearContents
    ws.Range("M7").Value2 = -100#: ws.Range("N7").Value2 = 3#
    ws.Range("M8").Value2 = 1000#: ws.Range("N8").Value2 = 3#
    RPX_SeepLastMsg = ""
End Sub
Public Function SeepUIFix3CancelCase(ByVal action As String) As Variant
    Dim caught As Long, message As String, k2 As Variant
    Call SeepUIFix3Prepare
    ' The public Stop action is also what a cancelled soil analysis leaves.
    Call RPX_Stop
    On Error Resume Next
    Select Case action
        Case "run": Call RPX_SeepRun
        Case "check": Call RPX_SeepCheckInputs
        Case "preview": Call RPX_SeepPreview
    End Select
    caught = Err.Number: message = Err.Description: Err.Clear
    k2 = RPX_Sheet("浸透結果").Range("K2").Value2
    SeepUIFix3CancelCase = Array(action, caught, message, RPX_SeepLastMsg, XCancel, RPX_Busy, RPX_InternalWrite, k2)
    On Error GoTo 0
End Function
Public Function SeepUIFix3FormulaCase() As Variant
    Dim ws As Worksheet, addresses As Variant, expected() As String, i As Long
    Dim before() As Variant, after() As Variant, lost As String, sourceBefore As Double, sourceAfter As Double
    Call SeepUIFix3Prepare
    Set ws = RPX_Sheet("浸透入力")
    addresses = Array("B3", "B4", "D4", "F4", "A7", "B7", "C7", "E7", "F7", "G7", "H7", "I7", "K7", "M7", "N7", "P7", "Q7", "C209")
    ReDim expected(0 To UBound(addresses)): ReDim before(0 To UBound(addresses)): ReDim after(0 To UBound(addresses))
    expected(0) = "=$Z$6": expected(1) = "=9.81": expected(2) = "=0": expected(3) = "=0"
    expected(4) = "=1": expected(5) = "=0.000001": expected(6) = "=0.000002"
    expected(7) = "=1": expected(8) = "=1": expected(9) = "=1": expected(10) = "=2": expected(11) = "=0": expected(12) = "=3"
    expected(13) = "=-100": expected(14) = "='材料データ'!B5/10": expected(15) = "=1": expected(16) = "=2": expected(17) = "=16"
    For i = 0 To UBound(addresses)
        ws.Range(addresses(i)).Formula = expected(i)
        before(i) = ws.Range(addresses(i)).Formula
    Next i
    ws.Calculate: sourceBefore = ws.Range("N7").Value2
    Call RPX_SeepSetupSheet
    For i = 0 To UBound(addresses)
        after(i) = ws.Range(addresses(i)).Formula
        If CStr(after(i)) <> CStr(before(i)) Then lost = lost & CStr(addresses(i)) & ","
    Next i
    RPX_Sheet("材料データ").Range("B5").Value2 = 51.65
    ws.Calculate: sourceAfter = ws.Range("N7").Value2
    SeepUIFix3FormulaCase = Array(lost, sourceBefore, sourceAfter, before, after, RPX_InternalWrite, RPX_SeepLastMsg)
End Function
Public Function SeepUIFix3LegacyFormulaCase() As Variant
    Dim ws As Worksheet, modeBefore As Long, modeAfter As Long
    Call SeepUIFix3Prepare
    Set ws = RPX_Sheet("浸透入力")
    ws.Cells.UnMerge: ws.Cells.Clear
    ws.Range("M6").Value2 = "項目"
    ws.Range("N24").Formula = "=IF('材料データ'!B5>0,3,2)"
    ws.Range("N13").Formula = "=9.81": ws.Range("N19").Formula = "=0": ws.Range("N20").Formula = "=0"
    ws.Range("N7").Formula = "=16"
    ws.Range("A7").Formula = "=1": ws.Range("B7").Formula = "=0.000001": ws.Range("C7").Formula = "=0.000002"
    ws.Range("W7").Formula = "=1": ws.Range("X7").Formula = "='材料データ'!C5"
    Call RPX_SeepSetupSheet
    modeBefore = RPX_SeepMode(ws)
    RPX_Sheet("材料データ").Range("B5").Value2 = -1
    RPX_Sheet("材料データ").Range("C5").Value2 = 16
    ws.Calculate: modeAfter = RPX_SeepMode(ws)
    SeepUIFix3LegacyFormulaCase = Array(modeBefore, modeAfter, ws.Range("B3").HasFormula, ws.Range("B4").Formula, ws.Range("C209").Formula, ws.Range("P7").Formula, ws.Range("Q7").HasFormula, ws.Range("Q7").Value2, RPX_InternalWrite, RPX_SeepLastMsg)
End Function
Public Function SeepUIFix3BusyFormatCase() As Variant
    Dim ws As Worksheet, oldFormula As String
    Call SeepUIFix3Prepare
    Set ws = RPX_Sheet("浸透入力")
    ws.Range("N7").Formula = "=3": oldFormula = ws.Range("N7").Formula
    RPX_Busy = True: RPX_InternalWrite = True: XCancel = True
    Call RPX_SeepSetupSheet
    SeepUIFix3BusyFormatCase = Array(oldFormula = ws.Range("N7").Formula, RPX_Busy, RPX_InternalWrite, XCancel, RPX_SeepLastMsg)
    RPX_Busy = False: RPX_InternalWrite = False: XCancel = False
End Function
