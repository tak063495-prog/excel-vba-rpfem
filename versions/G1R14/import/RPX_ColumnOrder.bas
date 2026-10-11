Attribute VB_Name = "RPX_ColumnOrder"
Option Explicit
' Independent general-LU column ordering on the implicit normal graph.
' Does not change the NT symmetric ordering or numeric operator.
' 1927正本と同じ挿入順・列挙中削除・一意キー・n>20000分岐。
' 変更点: 隣接は RPX_Row、近傍 Keys を吸収ループの前に写す。
' version=20260926_0630 engine=20260926_order_stamp
Private adjacency() As Object, selected() As Long
Private outputOrder() As Long, tag As Long, written As Long
Private heap() As Long, heapAt() As Long, heapSize As Long
Private priority() As Long
Private touchedReuse As Object
Private neighborMark() As Long, cliqueMark() As Long

Public Sub RPX_ColumnOrdering(ByVal n As Long, ByRef columns() As Object, ByRef permutation() As Long, ByRef inverse() As Long)
    Dim i As Long, j As Variant
    ReDim adjacency(0 To n - 1): ReDim selected(0 To n - 1)
    ReDim outputOrder(0 To n - 1)
    tag = 0: written = 0
    For i = 0 To n - 1
        Set adjacency(i) = RPX_Row()
    Next i
    QuotientDegree n, columns
    If written <> n Then err.Raise 5, , "ORDERING_INCOMPLETE"
    ReDim permutation(0 To n - 1): ReDim inverse(0 To n - 1)
    For i = 0 To n - 1
        permutation(i) = outputOrder(i): inverse(permutation(i)) = i
        RPX_ReleaseRow adjacency(i)
        Set adjacency(i) = Nothing
    Next i
End Sub

Private Sub HeapSwap(ByVal a As Long, ByVal b As Long)
    Dim v As Long
    v = heap(a): heap(a) = heap(b): heap(b) = v
    heapAt(heap(a)) = a: heapAt(heap(b)) = b
End Sub
Private Sub HeapUpdate(ByVal v As Long)
    Dim p As Long, j As Long, par As Long
    p = heapAt(v)
    Do While p > 0
        par = (p - 1) \ 2
        If priority(heap(par)) <= priority(v) Then Exit Do
        HeapSwap p, par: p = par
    Loop
    Do
        j = 2 * p + 1
        If j >= heapSize Then Exit Do
        If j + 1 < heapSize Then
            If priority(heap(j + 1)) < priority(heap(j)) Then j = j + 1
        End If
        If priority(v) <= priority(heap(j)) Then Exit Do
        HeapSwap p, j: p = j
    Loop
End Sub
Private Sub QuotientDegree(ByVal n As Long, ByRef normalRows() As Object)
    Dim incident() As Object, clique() As Object, neighbors As Object, touched As Object
    Dim externalDegree() As Long
    Dim i As Long, v As Long, j As Variant, k As Variant, el As Variant, degree As Long
    Dim nbr() As Long, nbrN As Long, raw As Variant, idx As Long, jj As Long
    ReDim incident(0 To n - 1): ReDim clique(0 To 2 * n - 1)
    Dim neighborEpoch As Long, cliqueEpoch As Long
    ReDim externalDegree(0 To 2 * n - 1)
    ReDim neighborMark(0 To n - 1): ReDim cliqueMark(0 To 2 * n - 1)
    ReDim heap(0 To n - 1): ReDim heapAt(0 To n - 1): ReDim priority(0 To n - 1): heapSize = n
    For i = 0 To n - 1
        Set incident(i) = RPX_Row(): heap(i) = i: heapAt(i) = i: priority(i) = adjacency(i).count
    Next i
    ' Each input row is an initial clique of the normal graph. No A^T A matrix.
    For i = 0 To n - 1
        Set clique(n + i) = RPX_Row()
        For Each j In normalRows(i).keys
            clique(n + i)(CLng(j)) = True: incident(CLng(j))(n + i) = True
        Next j
    Next i
    For i = 0 To n - 1
        tag = tag + 1: selected(i) = tag: degree = 0
        For Each el In incident(i).keys
            For Each k In clique(CLng(el)).keys
                If selected(CLng(k)) <> tag Then selected(CLng(k)) = tag: degree = degree + 1
            Next k
        Next el
        priority(i) = degree
    Next i
    For i = n - 1 To 0 Step -1: HeapUpdate heap(i): Next i
    Do While heapSize > 0
        If written Mod 8192 = 0 Then RPX_Stage "ordering " & written & "/" & n
        v = heap(0): Set neighbors = RPX_Row()
        For Each j In adjacency(v).keys: neighbors(CLng(j)) = True: Next j
        For Each el In incident(v).keys
            For Each j In clique(CLng(el)).keys: neighbors(CLng(j)) = True: Next j
        Next el
        If neighbors.Exists(v) Then neighbors.Remove v
        For Each el In incident(v).keys
            For Each j In clique(CLng(el)).keys: incident(CLng(j)).Remove CLng(el): Next j
            Set clique(CLng(el)) = Nothing
        Next el
        heapSize = heapSize - 1
        If heapSize > 0 Then heap(0) = heap(heapSize): heapAt(heap(0)) = 0: HeapUpdate heap(0)
        outputOrder(written) = v: written = written + 1
        nbrN = neighbors.count
        If nbrN > 0 Then
            raw = neighbors.keys
            ReDim nbr(0 To nbrN - 1)
            For idx = 0 To nbrN - 1
                nbr(idx) = CLng(raw(idx))
            Next idx
        End If
        neighborEpoch = neighborEpoch + 1
        If neighborEpoch <= 0 Then
            For i = 0 To n - 1: neighborMark(i) = 0: Next i
            neighborEpoch = 1
        End If
        For idx = 0 To nbrN - 1: neighborMark(nbr(idx)) = neighborEpoch: Next idx
        For idx = 0 To nbrN - 1
            jj = nbr(idx)
            If adjacency(jj).Exists(v) Then adjacency(jj).Remove v
            For Each k In adjacency(jj).keys
                If neighborMark(CLng(k)) = neighborEpoch Then adjacency(jj).Remove CLng(k)
            Next k
            incident(jj)(v) = True
        Next idx
        adjacency(v).RemoveAll: incident(v).RemoveAll: Set clique(v) = neighbors
        If n > 20000 Then
            For idx = 0 To nbrN - 1
                jj = nbr(idx)
                tag = tag + 1: selected(jj) = tag: degree = 0
                For Each k In adjacency(jj).keys
                    If selected(CLng(k)) <> tag Then selected(CLng(k)) = tag: degree = degree + 1
                Next k
                For Each el In incident(jj).keys
                    For Each k In clique(CLng(el)).keys
                        If selected(CLng(k)) <> tag Then selected(CLng(k)) = tag: degree = degree + 1
                    Next k
                Next el
                priority(jj) = degree: HeapUpdate jj
            Next idx
        Else
            cliqueEpoch = cliqueEpoch + 1
            If cliqueEpoch <= 0 Then
                For i = 0 To 2 * n - 1: cliqueMark(i) = 0: Next i
                cliqueEpoch = 1
            End If
            For idx = 0 To nbrN - 1
                jj = nbr(idx)
                For Each el In incident(jj).keys
                    If cliqueMark(CLng(el)) <> cliqueEpoch Then
                        cliqueMark(CLng(el)) = cliqueEpoch: degree = 0
                        For Each k In clique(CLng(el)).keys
                            If neighborMark(CLng(k)) <> neighborEpoch Then degree = degree + 1
                        Next k
                        externalDegree(CLng(el)) = degree
                    End If
                Next el
            Next idx
            For idx = 0 To nbrN - 1
                jj = nbr(idx)
                degree = adjacency(jj).count + nbrN - 1
                For Each el In incident(jj).keys
                    degree = degree + externalDegree(CLng(el))
                Next el
                If degree > heapSize - 1 Then degree = heapSize - 1
                priority(jj) = degree: HeapUpdate jj
            Next idx
        End If
    Loop
End Sub

