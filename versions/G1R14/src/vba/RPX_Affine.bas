Attribute VB_Name = "RPX_Affine"
Option Explicit
' Constructive substitution for fixed-load, zero-objective witnesses only.
' No numerical rank truncation. Original rows/cones remain authoritative.
Public AFOriginalN As Long, AFActive As Boolean
Public AFMap() As Long, AFWeight() As Double, AFOffset() As Double
Public AFRemovedRows As Collection, AFSweeps As Long
Private originalRows As Collection, originalRhs As Collection
Private originalIndexedRows() As Object, originalIndexedRhs() As Double
Private originalCones() As RPX_ConeData, originalXB As Long
Private originalModel As String, originalFs As String, originalQ As Long, originalEpoch As Long
Private parent() As Long, linkWeight() As Double, fixed() As Boolean, fixedValue() As Double
Private reducedId() As Long
Public AFMaxEquality As Double, AFMaxCone As Double
Public AFRawScale As Double, AFPhysicalGate As Boolean
Public Sub RPX_AffineOriginalContext(ByVal rows As Collection, ByVal rhs As Collection)
    AFActive = False: AFOriginalN = XN: originalXB = XB
    Set originalRows = rows: Set originalRhs = rhs:
    RPX_IndexRows rows, rhs, originalIndexedRows, originalIndexedRhs
     originalCones = XC
    originalFs = RPX_ExactDouble(RStrengthFactor): originalQ = RSubdivisions: originalEpoch = RPX_InputEpoch
    originalModel = ""
    If AFPhysicalGate Then originalModel = RPX_ExactModel()
End Sub
Public Sub RPX_AffineLoadMap(ByVal path As String)
    Dim f As Integer, n As Long, i As Long, number As Long, source As String, message As String
    On Error GoTo Failed
    f = FreeFile: Open path For Binary Access Read As #f
    Get #f, , n
    If n <> AFOriginalN Then err.Raise 9, "RPX_Affine", "AFFINE_MAP_DIMENSION"
    ReDim AFMap(0 To n - 1): ReDim AFWeight(0 To n - 1): ReDim AFOffset(0 To n - 1)
    For i = 0 To n - 1
        Get #f, , AFMap(i): Get #f, , AFWeight(i): Get #f, , AFOffset(i)
        If AFMap(i) < -1 Or AFMap(i) >= XN Then err.Raise 9, "RPX_Affine", "AFFINE_MAP_COLUMN"
        If Not RPX_SchurFinite(AFWeight(i)) Or Not RPX_SchurFinite(AFOffset(i)) Then err.Raise 5, "RPX_Affine", "AFFINE_MAP_NONFINITE"
    Next i
    If Seek(f) <> LOF(f) + 1 Then err.Raise 5, "RPX_Affine", "AFFINE_MAP_TRAILING_DATA"
    Close #f: AFActive = True: Exit Sub
Failed:
    number = err.number: source = err.source: message = err.description
    On Error Resume Next: Close #f: On Error GoTo 0
    AFActive = False: err.Raise number, source, message
End Sub
Public Sub RPX_AddExact(ByVal row As Object, ByVal id As Long, ByVal value As Double)
    If Not RPX_SchurFinite(value) Then err.Raise 5, "RPX_Affine", "AFFINE_NONFINITE_COEFFICIENT"
    If row.Exists(id) Then value = value + CDbl(row(id))
    If value = 0# Then
        If row.Exists(id) Then row.Remove id
    Else
        row(id) = value
    End If
End Sub
Private Function Root(ByVal id As Long, ByRef weight As Double) As Long
    Dim nextID As Long
    weight = 1#
    Do While parent(id) <> id
        nextID = parent(id): weight = weight * linkWeight(id): id = nextID
    Loop
    Root = id
End Function
Private Function Transform(ByVal row As Object, ByRef shift As Double) As Object
    Dim id As Variant, rootId As Long, weight As Double, value As Double
    Set Transform = RPX_Row(): shift = 0#
    For Each id In row.keys
        rootId = Root(CLng(id), weight): value = CDbl(row(id)) * weight
        If fixed(rootId) Then
            shift = shift + value * fixedValue(rootId)
        Else
            RPX_AddExact Transform, reducedId(rootId), value
        End If
    Next id
End Function
Public Sub RPX_AffinePrepare(ByRef rows As Collection, ByRef rhs As Collection)
    Dim indexedRows() As Object, indexedRhs() As Double
    Dim n As Long, i As Long, j As Long, k As Long, r As Long, sweep As Long, roots As Long
    Dim row As Object, keys As Variant, a As Long, b As Long, ra As Long, rb As Long
    Dim wa As Double, wb As Double, ca As Double, cb As Double, value As Double, shift As Double
    Dim consumed() As Boolean, changed As Boolean, extra As Boolean, allConstant As Boolean
    Dim nextRows As Collection, nextRhs As Collection, coneRows As Collection
    Dim nextCones() As RPX_ConeData, coneN As Long, coneOffsets() As Double
    Dim headEmpty As Boolean, norm As Double, oldMap As Long, oldWeight As Double
    AFActive = False: AFOriginalN = XN: originalXB = XB
    Set originalRows = rows: Set originalRhs = rhs
    RPX_IndexRows rows, rhs, originalIndexedRows, originalIndexedRhs
    originalCones = XC
    originalModel = "": originalFs = RPX_ExactDouble(RStrengthFactor)
    If AFPhysicalGate Then originalModel = RPX_ExactModel()
    originalQ = RSubdivisions: originalEpoch = RPX_InputEpoch
    NormalizeWorkingEqualities rows, rhs
    Set AFRemovedRows = New Collection
    ReDim AFMap(0 To AFOriginalN - 1): ReDim AFWeight(0 To AFOriginalN - 1): ReDim AFOffset(0 To AFOriginalN - 1)
    For i = 0 To AFOriginalN - 1: AFMap(i) = i: AFWeight(i) = 1#: Next i
    For sweep = 0 To 11
        RPX_AssertInputs "affine_prepare": RPX_Stage "affine substitution " & sweep
        RPX_IndexRows rows, rhs, indexedRows, indexedRhs
        n = XN
        ReDim parent(0 To n - 1): ReDim linkWeight(0 To n - 1)
        ReDim fixed(0 To n - 1): ReDim fixedValue(0 To n - 1): ReDim reducedId(0 To n - 1)
        ReDim consumed(1 To rows.count)
        For i = 0 To n - 1: parent(i) = i: linkWeight(i) = 1#: reducedId(i) = -1: Next i
        For i = 1 To rows.count
            Set row = indexedRows(i)
            If row.count = 0 Then
                If indexedRhs(i) <> 0# Then err.Raise 5, "RPX_Affine", "AFFINE_EMPTY_CONTRADICTION"
            ElseIf row.count = 1 Then
                keys = row.keys: a = CLng(keys(0)): ca = CDbl(row(a))
                ' Eligibility threshold is a pivot choice, never a row-deletion rule.
                If Abs(ca) >= 0.0000000001 Then
                    value = indexedRhs(i) / ca
                    If fixed(a) Then
                        If fixedValue(a) <> value Then err.Raise 5, "RPX_Affine", "AFFINE_SINGLETON_CONFLICT"
                    Else
                        fixed(a) = True: fixedValue(a) = value
                    End If
                    consumed(i) = True
                End If
            End If
        Next i
        For i = 1 To rows.count
            Set row = indexedRows(i)
            If row.count = 2 And indexedRhs(i) = 0# Then
                keys = row.keys: a = CLng(keys(0)): b = CLng(keys(1))
                ra = Root(a, wa): rb = Root(b, wb)
                If ra <> rb And Not fixed(ra) And Not fixed(rb) Then
                    ca = CDbl(row(a)) * wa: cb = CDbl(row(b)) * wb
                    If Abs(ca) >= 0.0000000001 Or Abs(cb) >= 0.0000000001 Then
                        If Abs(ca) >= 0.05 * Abs(cb) And Abs(cb) >= 0.05 * Abs(ca) Then
                            If Abs(ca) >= Abs(cb) Then
                                parent(ra) = rb: linkWeight(ra) = -cb / ca
                            Else
                                parent(rb) = ra: linkWeight(rb) = -ca / cb
                            End If
                            consumed(i) = True
                        End If
                    End If
                End If
            End If
        Next i
        roots = 0
        For i = 0 To n - 1
            If parent(i) = i And Not fixed(i) Then reducedId(i) = roots: roots = roots + 1
        Next i
        If roots < 1 Then err.Raise 5, "RPX_Affine", "AFFINE_NO_FREE_VARIABLES"
        changed = (roots < n)
        For i = 0 To AFOriginalN - 1
            oldMap = AFMap(i)
            If oldMap >= 0 Then
                r = Root(oldMap, wa): oldWeight = AFWeight(i)
                If fixed(r) Then
                    AFOffset(i) = AFOffset(i) + oldWeight * wa * fixedValue(r)
                    AFMap(i) = -1: AFWeight(i) = 0#
                Else
                    AFMap(i) = reducedId(r): AFWeight(i) = oldWeight * wa
                End If
            End If
        Next i
        Set nextRows = New Collection: Set nextRhs = New Collection
        For i = 1 To rows.count
            Set row = indexedRows(i)
            If consumed(i) And changed Then
                AFRemovedRows.Add "sweep=" & sweep & ";row=" & i
            Else
                Set row = Transform(row, shift): value = indexedRhs(i) - shift
                If row.count = 0 Then
                    If value <> 0# Then err.Raise 5, "RPX_Affine", "AFFINE_RESTORED_ROW_CONTRADICTION"
                Else
                    nextRows.Add row: nextRhs.Add value
                End If
            End If
        Next i
        nextCones = XC: coneN = XB: XB = 0: XS = 0: extra = False
        XN = roots
        Call TransformConesTyped(nextCones, coneN, nextRows, nextRhs, extra)
        Set rows = nextRows: Set rhs = nextRhs: AFSweeps = sweep + 1
        If Not changed And Not extra Then Exit For
    Next sweep
    If XB <= 0 Then err.Raise 5, "RPX_Affine", "AFFINE_NO_ACTIVE_CONES"
    Call equilibrate(rows, rhs)
    AFActive = True
    RPX_DiagEvent "affine_prepare", "original_n=" & AFOriginalN & ";reduced_n=" & XN & ";rows=" & rows.count & ";cones=" & XB & ";sweeps=" & AFSweeps & ";purpose=FIXED_FS_WITNESS"
End Sub
Private Sub TransformConesTyped(ByRef oldCones() As RPX_ConeData, ByVal coneN As Long, ByVal rows As Collection, ByVal rhs As Collection, ByRef extra As Boolean)
    Dim k As Long, j As Long, r As Long, i As Long, at As Long, q As Long, id As Long, rootId As Long
    Dim rowCount(0 To 2) As Long, rowId() As Long, rowValue() As Double, offsets(0 To 2) As Double
    Dim mapped() As Long, weight() As Double, fixedPart() As Double, unionIds() As Long, unionN As Long
    Dim maxCount As Long, n As Long, d As Long, value As Double, shift As Double, norm As Double
    Dim allConstant As Boolean, headEmpty As Boolean, row As Object
    For k = 0 To coneN - 1
        If oldCones(k).count > maxCount Then maxCount = oldCones(k).count
    Next k
    ReDim rowId(0 To 2, 0 To maxCount - 1): ReDim rowValue(0 To 2, 0 To maxCount - 1)
    ReDim mapped(0 To maxCount - 1): ReDim weight(0 To maxCount - 1): ReDim fixedPart(0 To maxCount - 1): ReDim unionIds(0 To 3 * maxCount - 1)
    For k = 0 To coneN - 1
        If k Mod 256 = 0 Then RPX_Stage "affine typed cones " & k & "/" & coneN
        n = oldCones(k).count: d = oldCones(k).dimn
        For j = 0 To n - 1
            rootId = Root(oldCones(k).ids(j), weight(j)): mapped(j) = reducedId(rootId): fixedPart(j) = 0#
            If fixed(rootId) Then mapped(j) = -1: fixedPart(j) = fixedValue(rootId)
        Next j
        allConstant = True
        For r = 0 To d - 1
            rowCount(r) = 0: shift = 0#
            For j = 0 To n - 1
                value = oldCones(k).coef(r, j)
                If value <> 0# Then
                    value = value * weight(j)
                    If mapped(j) < 0 Then
                        shift = shift + value * fixedPart(j)
                    Else
                        at = 0
                        Do While at < rowCount(r)
                            If rowId(r, at) = mapped(j) Then Exit Do
                            at = at + 1
                        Loop
                        If at < rowCount(r) Then value = rowValue(r, at) + value
                        If value = 0# Then
                            If at < rowCount(r) Then
                                For q = at To rowCount(r) - 2
                                    rowId(r, q) = rowId(r, q + 1): rowValue(r, q) = rowValue(r, q + 1)
                                Next q
                                rowCount(r) = rowCount(r) - 1
                            End If
                        Else
                            If at = rowCount(r) Then rowCount(r) = rowCount(r) + 1
                            rowId(r, at) = mapped(j): rowValue(r, at) = value
                        End If
                    End If
                End If
            Next j
            offsets(r) = oldCones(k).offset(r) + shift
            If rowCount(r) > 0 Then allConstant = False
        Next r
        headEmpty = (rowCount(0) = 0 And offsets(0) = 0#)
        If d > 1 And headEmpty Then
            For r = 1 To d - 1
                If rowCount(r) > 0 Then
                    Set row = RPX_Row()
                    For j = 0 To rowCount(r) - 1: RPX_AddExact row, rowId(r, j), rowValue(r, j): Next j
                    rows.Add row: rhs.Add -offsets(r): extra = True
                ElseIf offsets(r) <> 0# Then
                    err.Raise 5, "RPX_Affine", "AFFINE_ZERO_FACE_CONTRADICTION"
                End If
            Next r
        ElseIf allConstant Then
            norm = 0#
            For r = 1 To d - 1: norm = norm + offsets(r) * offsets(r): Next r
            If offsets(0) < Sqr(norm) Then err.Raise 5, "RPX_Affine", "AFFINE_CONSTANT_CONE_CONTRADICTION"
        Else
            unionN = 0
            For r = 0 To d - 1
                For j = 0 To rowCount(r) - 1
                    id = rowId(r, j): at = 0
                    Do While at < unionN
                        If unionIds(at) = id Then Exit Do
                        at = at + 1
                    Loop
                    If at = unionN Then unionIds(unionN) = id: unionN = unionN + 1
                Next j
            Next r
            With XC(XB)
                .dimn = d: .count = unionN: .first = XS: XS = XS + d
                .modelKind = 0: .owner = 0: .parameter1 = 0: .parameter2 = 0: .directionSign = 0
                ReDim .ids(0 To unionN - 1): ReDim .coef(0 To d - 1, 0 To unionN - 1): ReDim .offset(0 To d - 1): ReDim .slots(0 To unionN - 1, 0 To unionN - 1)
                For j = 0 To unionN - 1: .ids(j) = unionIds(j): Next j
                For r = 0 To d - 1
                    .offset(r) = offsets(r)
                    For j = 0 To rowCount(r) - 1
                        For at = 0 To unionN - 1
                            If unionIds(at) = rowId(r, j) Then .coef(r, at) = rowValue(r, j): Exit For
                        Next at
                    Next j
                Next r
            End With
            XB = XB + 1
        End If
    Next k
End Sub
Private Function RuizMultiplier(ByVal norm As Double) As Double
    If norm < 0.0001 Then norm = 0.0001
    If norm > 10000# Then norm = 10000#
    RuizMultiplier = 1# / Sqr(norm)
End Function
Private Sub equilibrate(ByVal rows As Collection, ByRef rhs As Collection)
    Dim indexedRows() As Object, indexedRhs() As Double
    Dim pass As Long, i As Long, j As Long, r As Long, key As Variant, row As Object
    Dim norm As Double, factor As Double, colNorm() As Double, colScale() As Double
    Dim nextRhs As Collection
    ' Data scales preserve each SOC block. KKT/NT scaling stays separate.
    RPX_IndexRows rows, rhs, indexedRows, indexedRhs
    For pass = 1 To 10
        ReDim colNorm(0 To XN - 1): ReDim colScale(0 To XN - 1)
        
        For i = 1 To rows.count
            Set row = indexedRows(i): norm = 0#
            For Each key In row.keys
                If Abs(CDbl(row(key))) > norm Then norm = Abs(CDbl(row(key)))
            Next key
            factor = RuizMultiplier(norm)
            For Each key In row.keys
                row(key) = CDbl(row(key)) * factor
                If Abs(CDbl(row(key))) > colNorm(CLng(key)) Then colNorm(CLng(key)) = Abs(CDbl(row(key)))
            Next key
            indexedRhs(i) = indexedRhs(i) * factor
        Next i
        
        For i = 0 To XB - 1
            norm = 0#
            For r = 0 To XC(i).dimn - 1
                For j = 0 To XC(i).count - 1
                    If Abs(XC(i).coef(r, j)) > norm Then norm = Abs(XC(i).coef(r, j))
                Next j
            Next r
            factor = RuizMultiplier(norm)
            For r = 0 To XC(i).dimn - 1
                XC(i).offset(r) = XC(i).offset(r) * factor
                For j = 0 To XC(i).count - 1
                    XC(i).coef(r, j) = XC(i).coef(r, j) * factor
                    If Abs(XC(i).coef(r, j)) > colNorm(XC(i).ids(j)) Then colNorm(XC(i).ids(j)) = Abs(XC(i).coef(r, j))
                Next j
            Next r
        Next i
        For j = 0 To XN - 1: colScale(j) = RuizMultiplier(colNorm(j)): Next j
        For i = 1 To rows.count
            Set row = indexedRows(i)
            For Each key In row.keys: row(key) = CDbl(row(key)) * colScale(CLng(key)): Next key
        Next i
        For i = 0 To XB - 1
            For r = 0 To XC(i).dimn - 1
                For j = 0 To XC(i).count - 1: XC(i).coef(r, j) = XC(i).coef(r, j) * colScale(XC(i).ids(j)): Next j
            Next r
        Next i
        For i = 0 To AFOriginalN - 1
            If AFMap(i) >= 0 Then AFWeight(i) = AFWeight(i) * colScale(AFMap(i))
        Next i
        RPX_Stage "affine equivalent scaling " & pass
    Next pass
    Set nextRhs = New Collection
    For i = 1 To rows.count: nextRhs.Add indexedRhs(i): Next i
    Set rhs = nextRhs
End Sub
Public Sub RPX_AffineRestoreX(ByRef reduced() As Double, ByRef restored() As Double)
    Dim i As Long
    If Not AFActive Then err.Raise 5, "RPX_Affine", "AFFINE_CONTEXT_MISSING"
    ReDim restored(0 To AFOriginalN - 1)
    For i = 0 To AFOriginalN - 1
        restored(i) = AFOffset(i)
        If AFMap(i) >= 0 Then restored(i) = restored(i) + AFWeight(i) * reduced(AFMap(i))
        If Not RPX_SchurFinite(restored(i)) Then err.Raise 5, "RPX_Affine", "AFFINE_NONFINITE_RESTORE"
    Next i
End Sub
Public Function RPX_AffineOriginalGate(ByRef reduced() As Double, Optional ByVal rawScale As Double = 1#) As Boolean
    Dim field() As Double, i As Long, j As Long, r As Long, key As Variant, row As Object
    Dim residual As Double, total As Double, correction As Double, term As Double, nextValue As Double
    Dim c(0 To 2) As Double, cone As Double
    If originalEpoch <> RPX_InputEpoch Or originalQ <> RSubdivisions Or originalFs <> RPX_ExactDouble(RStrengthFactor) Then err.Raise 5, "RPX_Affine", "AFFINE_CONTEXT_MISMATCH"
    RPX_AssertInputs "affine_original_gate"
    If AFPhysicalGate Then
        If Len(originalModel) = 0 Then err.Raise 5, "RPX_Affine", "AFFINE_PHYSICAL_IDENTITY_MISSING"
        If originalModel <> RPX_ExactModel() Then err.Raise 5, "RPX_Affine", "AFFINE_PHYSICAL_IDENTITY_MISMATCH"
    End If
    RPX_AffineRestoreX reduced, field
    AFMaxEquality = 0#: AFMaxCone = 0#
    For i = 1 To originalRows.count
        Set row = originalIndexedRows(i): total = 0#: correction = 0#
        For Each key In row.keys
            term = CDbl(row(key)) * field(CLng(key)) - correction
            nextValue = total + term: correction = (nextValue - total) - term: total = nextValue
        Next key
        residual = Abs(total - originalIndexedRhs(i))
        If residual > AFMaxEquality Then AFMaxEquality = residual
    Next i
    For i = 0 To originalXB - 1
        For r = 0 To originalCones(i).dimn - 1
            total = originalCones(i).offset(r): correction = 0#
            For j = 0 To originalCones(i).count - 1
                term = originalCones(i).coef(r, j) * field(originalCones(i).ids(j)) - correction
                nextValue = total + term: correction = (nextValue - total) - term: total = nextValue
            Next j
            c(r) = total
        Next r
        cone = -c(0)
        If originalCones(i).dimn = 3 Then cone = Sqr(c(1) * c(1) + c(2) * c(2)) - c(0)
        If cone > AFMaxCone Then AFMaxCone = cone
    Next i
    RPX_AffineOriginalGate = (AFMaxEquality <= 0.0000001 And AFMaxCone <= 0.0000001 And AFMaxEquality * rawScale <= 0.0000001 And AFMaxCone * rawScale <= 0.0000001)
End Function
Public Sub RPX_AffineEnd()
    AFActive = False
    Set originalRows = Nothing: Set originalRhs = Nothing: Set AFRemovedRows = Nothing
    Erase originalIndexedRows: Erase originalIndexedRhs
    Erase originalCones: Erase AFMap: Erase AFWeight: Erase AFOffset
End Sub

Public Function RPX_AffineCandidate(ByRef candidate As RPX_HsdCandidateState) As Boolean
    Dim restored() As Double, oldX() As Double, oldValue As Double, oldAudit As Double
    Dim raw(1 To 4) As Double, i As Long, number As Long, source As String, message As String, accepted As Boolean
    Dim oldGate As Boolean, oldReason As String
    If Not RPX_AffineOriginalGate(candidate.x, AFRawScale) Then Exit Function
    If AFPhysicalGate Then
        RPX_AffineRestoreX candidate.x, restored
        oldX = xx: oldValue = RValue: oldAudit = RAudit
        oldGate = RPhysicalGateAccepted: oldReason = RPhysicalGateReason
        For i = 1 To 4: raw(i) = RRawLowerAudit(i): Next i
        On Error GoTo Failed
        xx = restored: Call RPX_AuditLower(True)
        accepted = RPhysicalGateAccepted
Restore:
        xx = oldX: RValue = oldValue: RAudit = oldAudit
        RPhysicalGateAccepted = oldGate: RPhysicalGateReason = oldReason
        For i = 1 To 4: RRawLowerAudit(i) = raw(i): Next i
        On Error GoTo 0
        If number <> 0 Then err.Raise number, source, message
        If Not accepted Then Exit Function
    End If
    RPX_AffineCandidate = True: Exit Function
Failed:
    number = err.number: source = err.source: message = err.description
    Resume Restore
End Function

Private Sub NormalizeWorkingEqualities(ByRef rows As Collection, ByRef rhs As Collection)
    Dim indexed() As Object, values() As Double, rebuiltRows As Collection, rebuiltRhs As Collection
    Dim inputRow As Object
    Dim i As Long, key As Variant, row As Object, largest As Double, magnitude As Double, rowMultiplier As Double, value As Double
    RPX_IndexRows rows, rhs, indexed, values
    Set rebuiltRows = New Collection: Set rebuiltRhs = New Collection
    For i = 1 To rows.count
        Set inputRow = indexed(i)
        largest = 0#
        For Each key In inputRow.keys
            value = Abs(CDbl(inputRow(key))): If value > largest Then largest = value
        Next key
        rowMultiplier = 1#: magnitude = largest
        If largest >= 1E-150 And largest <= 1E+150 Then
            Do While magnitude >= 2#: magnitude = magnitude * 0.5: rowMultiplier = rowMultiplier * 0.5: Loop
            Do While magnitude < 1#: magnitude = magnitude * 2#: rowMultiplier = rowMultiplier * 2#: Loop
        End If
        Set row = RPX_Row()
        For Each key In inputRow.keys
            value = CDbl(inputRow(key)) * rowMultiplier
            If value = 0# And CDbl(inputRow(key)) <> 0# Then err.Raise 5, "RPX_Affine", "AFFINE_ROW_SCALE_LOSS"
            RPX_AddExact row, CLng(key), value
        Next key
        value = values(i) * rowMultiplier
        If value = 0# And values(i) <> 0# Then err.Raise 5, "RPX_Affine", "AFFINE_RHS_SCALE_LOSS"
        rebuiltRows.Add row: rebuiltRhs.Add value
    Next i
    Set rows = rebuiltRows: Set rhs = rebuiltRhs
End Sub
