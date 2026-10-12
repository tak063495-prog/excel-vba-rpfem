Attribute VB_Name = "ThisWorkbook"
Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = True

Private Sub Workbook_SheetChange(ByVal sh As Object, ByVal target As Range)
    If RPX_InternalWrite Then Exit Sub
    If RPX_IsInputChange(sh, target) Then RPX_Invalidate
End Sub

Private Sub Workbook_SheetCalculate(ByVal sh As Object)
    If RPX_InternalWrite Or RPX_Busy Then Exit Sub
    If RPX_InputSheetIndex(CStr(sh.name)) >= 0 Then
        If RPX_InputsChanged() Then RPX_Invalidate
    End If
End Sub

