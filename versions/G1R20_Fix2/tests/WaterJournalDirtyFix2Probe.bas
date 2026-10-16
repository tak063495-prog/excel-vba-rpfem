Attribute VB_Name = "WaterJournalDirtyFix2Probe"
Option Explicit

Public Function WaterJournalDirtyFix2() As Variant
    Dim savedBefore As Boolean, savedAfter As Boolean, k2 As Variant, k9 As Variant, result As Worksheet
    Dim item As Object, marker As String
    Set result = RPX_Sheet("浸透結果")
    k2 = result.Range("K2").Value2: k9 = result.Range("K9").Value2
    ThisWorkbook.Save
    savedBefore = ThisWorkbook.Saved
    ' No worksheet operation occurs in this public gate. Only the journal and
    ' in-memory interpolation arrays change between the two Saved readings.
    Call RPX_WaterResultWriteBegin
    savedAfter = ThisWorkbook.Saved
    For Each item In ThisWorkbook.CustomDocumentProperties
        If CStr(item.Name) = "RPX_SeepResultWriteState" Then marker = CStr(item.Value)
    Next item
    If CStr(k2) <> CStr(result.Range("K2").Value2) Or CStr(k9) <> CStr(result.Range("K9").Value2) Then Err.Raise 5, , "RESULT_SHEET_CHANGED"
    WaterJournalDirtyFix2 = Array("OBSERVED", savedBefore, savedAfter, marker, "result_K2_K9_unchanged")
End Function
