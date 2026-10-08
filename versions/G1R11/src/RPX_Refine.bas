Attribute VB_Name = "RPX_Refine"
Option Explicit
' Conforming subdivision, preserving every old triangle and material interface.
' The old P2 fields restrict to P2 fields on the subtriangles.
Public Sub RPX_RefineMesh(ByRef indicator() As Double, ByVal fraction As Double, Optional ByVal forceLegacy As Boolean = False)
    Dim marks As Object, key As Variant, rec As Variant, rank() As Long, i As Long, j As Long
    Dim count As Long, e As Long, k As Long, a As Long, b As Long, center As Long, oldRE As Long
    Dim oldTri() As Long, poly(0 To 5) As Long, np As Long, hit As Boolean, nodes As Long
    Dim red As Boolean, mids(0 To 2) As Long, marked As Long, a0 As Long, b0 As Long, c0 As Long, ab As Long, ca As Long
    red = RPX_RobustMeshPolicy() And Not forceLegacy
    If fraction <= 0# Or fraction > 1# Then err.Raise 5, , "REFINE_FRACTION_OUT_OF_RANGE"
    RProgramKind = "": oldRE = re: oldTri = RTri
    RefineRank indicator, rank
    count = CLng(re * fraction): If count < 1 Then count = 1
    Set marks = RPX_Row()
    For i = 0 To count - 1
        e = rank(i)
        For k = 0 To 2
            a = RTri(e, k): b = RTri(e, (k + 1) Mod 3)
            key = RPX_EdgeKey(a, b)
            If Not marks.Exists(key) Then
                If rn >= UBound(RXY, 1) - 3 Then err.Raise 7, , "REFINE_NODE_CAPACITY"
                marks(key) = rn
                RXY(rn, 0) = (RXY(a, 0) + RXY(b, 0)) / 2#: RXY(rn, 1) = (RXY(a, 1) + RXY(b, 1)) / 2#
                If RConstraints.Exists(key) Then
                    RConstraints.Remove key
                    RConstraints(RPX_EdgeKey(a, rn)) = Array(a, rn)
                    RConstraints(RPX_EdgeKey(rn, b)) = Array(rn, b)
                End If
                rn = rn + 1
            End If
        Next k
    Next i
    re = 0
    For e = 0 To oldRE - 1
        np = 0: hit = False: marked = 0
        For k = 0 To 2: mids(k) = -1: Next k
        For k = 0 To 2
            a = oldTri(e, k): b = oldTri(e, (k + 1) Mod 3)
            poly(np) = a: np = np + 1: key = RPX_EdgeKey(a, b)
            If marks.Exists(key) Then poly(np) = marks(key): mids(k) = marks(key): np = np + 1: hit = True: marked = marked + 1
        Next k
        If Not hit Then
            For k = 0 To 2: RTri(re, k) = oldTri(e, k): Next k
            re = re + 1
        ElseIf red Then
            If marked = 3 Then
                RefineChild oldTri(e, 0), mids(0), mids(2)
                RefineChild mids(0), oldTri(e, 1), mids(1)
                RefineChild mids(2), mids(1), oldTri(e, 2)
                RefineChild mids(0), mids(1), mids(2)
            ElseIf marked = 1 Then
                For k = 0 To 2: If mids(k) >= 0 Then Exit For
                Next k
                a0 = oldTri(e, k): b0 = oldTri(e, (k + 1) Mod 3): c0 = oldTri(e, (k + 2) Mod 3)
                RefineChild a0, mids(k), c0
                RefineChild mids(k), b0, c0
            Else
                For k = 0 To 2: If mids(k) >= 0 And mids((k + 2) Mod 3) >= 0 Then Exit For
                Next k
                a0 = oldTri(e, k): b0 = oldTri(e, (k + 1) Mod 3): c0 = oldTri(e, (k + 2) Mod 3)
                ab = mids(k): ca = mids((k + 2) Mod 3)
                RefineChild a0, ab, ca
                If (RXY(ab, 0) - RXY(c0, 0)) ^ 2 + (RXY(ab, 1) - RXY(c0, 1)) ^ 2 <= (RXY(ca, 0) - RXY(b0, 0)) ^ 2 + (RXY(ca, 1) - RXY(b0, 1)) ^ 2 Then
                    RefineChild ab, b0, c0: RefineChild ab, c0, ca
                Else
                    RefineChild ab, b0, ca: RefineChild b0, c0, ca
                End If
            End If
        Else
            center = rn
            If rn >= UBound(RXY, 1) - 3 Then err.Raise 7, , "REFINE_NODE_CAPACITY"
            RXY(rn, 0) = 0#: RXY(rn, 1) = 0#
            For k = 0 To 2
                RXY(rn, 0) = RXY(rn, 0) + RXY(oldTri(e, k), 0) / 3#
                RXY(rn, 1) = RXY(rn, 1) + RXY(oldTri(e, k), 1) / 3#
            Next k
            rn = rn + 1
            For k = 0 To np - 1
                If re > UBound(RTri, 1) Then err.Raise 7, , "REFINE_ELEMENT_CAPACITY"
                RPX_SetPositive re, poly(k), poly((k + 1) Mod np), center: re = re + 1
            Next k
        End If
    Next e
    Call RPX_AssignRegions
End Sub
Public Function RPX_TestRefine() As String
    Dim eta() As Double, e As Long, area As Double, oldElements As Long
    RPX_SlopeDefinition 64: Call RPX_GenerateMesh
    oldElements = re: ReDim eta(0 To re - 1)
    For e = 0 To re - 1: eta(e) = e + 1#: Next e
    RPX_RefineMesh eta, 0.2
    For e = 0 To re - 1: area = area + RArea(e): Next e
    If Abs(area - 1650#) > 0.000001 Or re <= oldElements Then err.Raise 5, , "REFINE_AREA_FAILED"
    For e = 0 To RNEdge - 1
        If REdges(e).element(1) < 0 Then
            If Not RConstraints.Exists(RPX_EdgeKey(REdges(e).a, REdges(e).b)) Then err.Raise 5, , "REFINE_HANGING_EDGE"
        End If
    Next e
    RPX_TestRefine = "PASS area=" & CStr(area) & " elements=" & CStr(oldElements) & " -> " & CStr(re)
End Function

Public Function RPX_RobustMeshPolicy() As Boolean
    Dim policy As String
    policy = UCase$(Trim$(CStr(RPX_Setting("ADAPT_MESH_POLICY", "LEGACY"))))
    Select Case policy
        Case "LEGACY"
        Case "ROBUST_REFINE", "LEGACY_THEN_REFINE": RPX_RobustMeshPolicy = (RR = 0)
        Case Else: err.Raise 5, "RPX_Refine", "INVALID_ADAPT_MESH_POLICY"
    End Select
End Function
Private Sub RefineChild(ByVal a As Long, ByVal b As Long, ByVal c As Long)
    If re > UBound(RTri, 1) Then err.Raise 7, , "REFINE_ELEMENT_CAPACITY"
    RPX_SetPositive re, a, b, c: re = re + 1
End Sub
Private Sub RefineRank(ByRef indicator() As Double, ByRef rank() As Long)
    Dim e As Long, i As Long, j As Long
    ReDim rank(0 To re - 1)
    For e = 0 To re - 1: rank(e) = e: Next e
    ' Stable insertion sort is sufficient for the initial 1536-element density.
    For i = 1 To re - 1
        e = rank(i): j = i - 1
        Do While j >= 0
            If indicator(rank(j)) >= indicator(e) Then Exit Do
            rank(j + 1) = rank(j): j = j - 1
        Loop
        rank(j + 1) = e
    Next i
End Sub
Public Function RPX_RefineBudgetFraction(ByRef indicator() As Double, ByVal cap As Long, ByVal maximum As Double) As Double
    Dim rank() As Long, costs As Object, marks As Object, key As String, e As Long, k As Long, i As Long
    Dim count As Long, selected As Long, needed As Long, elements As Long, a As Long, b As Long
    If cap <= re Then Exit Function
    RefineRank indicator, rank
    Set costs = RPX_Row(): Set marks = RPX_Row()
    For i = 0 To RNEdge - 1
        key = RPX_EdgeKey(REdges(i).a, REdges(i).b)
        If REdges(i).element(1) < 0 Then costs(key) = 1 Else costs(key) = 2
    Next i
    count = CLng(re * maximum): If count < 1 Then count = 1
    elements = re
    For i = 0 To count - 1
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        e = rank(i): needed = 0
        For k = 0 To 2
            a = RTri(e, k): b = RTri(e, (k + 1) Mod 3): key = RPX_EdgeKey(a, b)
            If Not marks.Exists(key) Then needed = needed + CLng(costs(key))
        Next k
        If elements + needed > cap Then Exit For
        For k = 0 To 2: key = RPX_EdgeKey(RTri(e, k), RTri(e, (k + 1) Mod 3)): marks(key) = 1: Next k
        elements = elements + needed: selected = selected + 1
    Next i
    If selected > 0 Then
        RPX_RefineBudgetFraction = selected / CDbl(re)
        If CLng(re * RPX_RefineBudgetFraction) <> selected Then err.Raise 5, , "REFINE_BUDGET_COUNT_MISMATCH"
    End If
    RPX_DiagEvent "adapt_refine_budget", "cap=" & cap & ";predicted_elements=" & elements & ";selected=" & selected
End Function
