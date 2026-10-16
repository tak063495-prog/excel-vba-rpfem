Attribute VB_Name = "InputMatrixFix4Probe"
Option Explicit

Public Function InputMatrixCase(ByVal test As String) As Variant
    Dim geometry As Worksheet, materials As Worksheet, caught As Long, message As String, i As Long, load As Double, shift As Double
    Call RPX_InputEnd: XCancel = False
    PhaseSeedInputFixture
    Set geometry = ThisWorkbook.Worksheets("要素定義")
    Set materials = ThisWorkbook.Worksheets("材料データ")
    Select Case test
        Case "valid"
        Case "c0": materials.Range("B5").Value2 = 0#
        Case "gamma0": materials.Range("D5").Value2 = 0#
        Case "davis": materials.Range("C5").Value2 = 45#: materials.Range("E5").Value2 = 30#
        Case "negative_coordinates"
            For i = 7 To 10
                geometry.Cells(i, 2).Value2 = geometry.Cells(i, 2).Value2 - 3#
                geometry.Cells(i, 3).Value2 = geometry.Cells(i, 3).Value2 - 2#
            Next i
        Case "numeric_text_id": geometry.Range("A7").NumberFormat = "@": geometry.Range("A7").Value2 = "1"
        Case "last_point_row": geometry.Range("A1006:C1006").Value2 = Array(5, 0.5, 0.5)
        Case "last_material_row": materials.Range("A1004:D1004").Value2 = Array(2, 10#, 0#, 0#)
        Case "overflow_point": geometry.Range("A1007:C1007").Value2 = Array(5, 0.5, 0.5)
        Case "overflow_material": materials.Range("A1005:D1005").Value2 = Array(2, 10#, 0#, 0#)
        Case "overflow_rigid": materials.Range("F1005:Q1005").Value2 = Array(1, 0.5, 0.5, 1, 1, 1, 0#, 0#, 0#, 1#, 0#, 0#)
        Case "overflow_region": geometry.Range("E1007:H1007").Value2 = Array(2, 1, 0, "1,2,3,4")
        Case "overflow_face": geometry.Range("J1007:T1007").Value2 = Array(4, 1, 2, 3, 0, "自由", 0, 0#, 0#, 0#, -2#)
        Case "adjacent_regions", "contained_regions", "contained_touching_region", "aligned_overlap", "aligned_overlap_narrow", "aligned_overlap_left"
            materials.Range("A6:D6").Value2 = Array(2, 50#, 20#, 21#)
            geometry.Range("E8:H8").Value2 = Array(2, 2, 0, "5,6,7,8")
            If test = "adjacent_regions" Then
                geometry.Range("A11:C11").Value2 = Array(5, 2#, 1#)
                geometry.Range("A12:C12").Value2 = Array(6, 2#, 0#)
                geometry.Range("H8").Value2 = "4,3,5,6"
            ElseIf test = "contained_regions" Then
                geometry.Range("A11:C11").Value2 = Array(5, 0.25, 0.25)
                geometry.Range("A12:C12").Value2 = Array(6, 0.25, 0.75)
                geometry.Range("A13:C13").Value2 = Array(7, 0.75, 0.75)
                geometry.Range("A14:C14").Value2 = Array(8, 0.75, 0.25)
            ElseIf test = "contained_touching_region" Then
                geometry.Range("A11:C11").Value2 = Array(5, 0.25, 0#)
                geometry.Range("A12:C12").Value2 = Array(6, 0.25, 1#)
                geometry.Range("A13:C13").Value2 = Array(7, 0.75, 1#)
                geometry.Range("A14:C14").Value2 = Array(8, 0.75, 0#)
            Else
                shift = 0.5
                If test = "aligned_overlap_narrow" Then shift = 0.999
                If test = "aligned_overlap_left" Then shift = -0.999
                geometry.Range("A11:C11").Value2 = Array(5, shift, 0#)
                geometry.Range("A12:C12").Value2 = Array(6, shift, 1#)
                geometry.Range("A13:C13").Value2 = Array(7, shift + 1#, 1#)
                geometry.Range("A14:C14").Value2 = Array(8, shift + 1#, 0#)
            End If
        Case "duplicate_point": geometry.Range("A8").Value2 = 1
        Case "zero_point": geometry.Range("A8").Value2 = 0
        Case "negative_point": geometry.Range("A8").Value2 = -1
        Case "fraction_point": geometry.Range("A8").Value2 = 1.5
        Case "boolean_point": geometry.Range("A8").Value2 = True
        Case "text_point": geometry.Range("A8").Value2 = "not an ID"
        Case "error_point": geometry.Range("A8").Value2 = CVErr(xlErrNA)
        Case "missing_referenced_point": geometry.Range("A8").ClearContents
        Case "duplicate_material": materials.Range("A6:D6").Value2 = Array(1, 10#, 0#, 0#)
        Case "negative_c": materials.Range("B5").Value2 = -1#
        Case "negative_gamma": materials.Range("D5").Value2 = -1#
        Case "phi89": materials.Range("C5").Value2 = 89#
        Case "psi_above_phi": materials.Range("E5").Value2 = 1#
        Case "error_c": materials.Range("B5").Value2 = CVErr(xlErrNum)
        Case "missing_material_ref": geometry.Range("F7").Value2 = 999
        Case "missing_rigid_ref": geometry.Range("G7").Value2 = 999
        Case "duplicate_region": geometry.Range("E8:H8").Value2 = Array(1, 1, 0, "1,2,3,4")
        Case "identical_regions": geometry.Range("E8:H8").Value2 = Array(2, 1, 0, "1,2,3,4")
        Case "repeated_polygon_point": geometry.Range("H7").Value2 = "1,2,3,2,4"
        Case "self_cross_polygon": geometry.Range("H7").Value2 = "1,3,2,4"
        Case "nonzero_self_cross_polygon"
            geometry.Range("A11:C11").Value2 = Array(5, 2#, 0.5)
            geometry.Range("H7").Value2 = "1,3,2,4,5"
        Case "duplicate_face": geometry.Range("J9").Value2 = 1
        Case "missing_face_region": geometry.Range("K9").Value2 = 999
        Case "missing_face_point": geometry.Range("L9").Value2 = 999
        Case "wrap_mismatch": geometry.Range("N9").Value2 = 1
        Case "unknown_restraint": geometry.Range("O9").Value2 = "invalid"
        Case "rigid_on_load_face": geometry.Range("P9").Value2 = 1
        Case "boolean_load": geometry.Range("T9").Value2 = True
        Case "boolean_x": geometry.Range("B7").Value2 = True
        Case "boolean_c_false": materials.Range("B5").Value2 = False
        Case "boolean_gamma_false": materials.Range("D5").Value2 = False
        Case "boolean_psi_false": materials.Range("E5").Value2 = False
        Case "error_load": geometry.Range("T9").Value2 = CVErr(xlErrValue)
        Case "partial_point": geometry.Range("B11:C11").Value2 = Array(2#, 3#)
        Case "partial_material": materials.Range("B6:D6").Value2 = Array(10#, 0#, 20#)
        Case "partial_rigid": materials.Range("G6:Q6").Value2 = Array(0.5, 0.5, 1, 1, 1, 0#, 0#, 0#, 1#, 0#, 0#)
        Case "partial_region": geometry.Range("F8:H8").Value2 = Array(1, 0, "1,2,3,4")
        Case "partial_face": geometry.Range("J9").ClearContents
        Case Else: Err.Raise 5, "InputMatrixFix4Probe", "UNKNOWN_CASE"
    End Select
    On Error GoTo Failed
    RPX_ReadInputs
    For i = 0 To DFCount - 1: load = load + DFaces(i).load1(1): Next i
    InputMatrixCase = Array("ACCEPTED", 0, "", RM, RR, DPCount, DRCount, DFCount, load, materials.Range("C5").Value2, materials.Range("E5").Value2, RMaterials(0).friction, RMaterials(0).dilation)
    Exit Function
Failed:
    caught = Err.Number: message = Err.Description
    InputMatrixCase = Array("REJECTED", caught, message, RM, RR, DPCount, DRCount, DFCount, Empty, materials.Range("C5").Value2, materials.Range("E5").Value2, RMaterials(0).friction, RMaterials(0).dilation)
End Function

Public Function LegacyBusyCase(ByVal path As String) As Variant
    Dim caught As Long, message As String, oldBusy As Boolean, current As Boolean, mesh As Boolean
    Call RPX_InputEnd: XCancel = False
    oldBusy = RPX_Busy
    RPX_ResultCurrent = True: RPX_MeshCurrent = True: RPX_Busy = True
    On Error Resume Next
    Call RPX_ReadModel(path)
    caught = Err.Number: message = Err.Description: Err.Clear
    On Error GoTo 0
    current = RPX_ResultCurrent: mesh = RPX_MeshCurrent
    RPX_Busy = oldBusy
    LegacyBusyCase = Array(caught, message, current, mesh)
End Function

Public Function LegacyInputCase(ByVal path As String) As Variant
    Dim number As Long, message As String, file As Integer, lockError As Long, accepted As Boolean
    Call RPX_InputEnd: XCancel = False
    ' State sentinels only; no field/root/solver status or certificate is supplied.
    RPX_ResultCurrent = True: RPX_MeshCurrent = True
    On Error GoTo ReadFailed
    Call RPX_ReadModel(path)
    accepted = True
    GoTo CheckClosed
ReadFailed:
    number = Err.Number: message = Err.Description
    Resume CheckClosed
CheckClosed:
    On Error Resume Next
    file = FreeFile
    Open path For Binary Access Read Write Lock Read Write As #file
    lockError = Err.Number: Err.Clear
    If lockError = 0 Then Close #file
    On Error GoTo 0
    If accepted Then
        LegacyInputCase = Array("ACCEPTED", 0, "", lockError, rn, re, RM, RTri(0, 0), RTri(0, 1), RTri(0, 2), RMaterials(0).cohesion, RMaterials(0).friction, RMaterials(0).gamma, RPX_ResultCurrent, RPX_MeshCurrent)
    Else
        LegacyInputCase = Array("REJECTED", number, message, lockError, rn, re, RM, Empty, Empty, Empty, Empty, Empty, Empty, RPX_ResultCurrent, RPX_MeshCurrent)
    End If
    ' Disposal-only probe cleanup after measuring the producer's file ownership.
    Close
End Function
