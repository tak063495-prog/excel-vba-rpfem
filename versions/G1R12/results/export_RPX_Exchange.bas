Option Explicit
' Explicit verification entry points. No Python dependency in the workbook.
Public Sub RPX_ExchangeMagic(ByVal f As Integer, ByVal text As String)
    Dim i As Long, b As Byte
    For i = 1 To Len(text): b = Asc(mid$(text, i, 1)): Put #f, , b: Next i
End Sub
Public Sub RPX_ExportCompiled(ByVal path As String)
    Dim f As Integer, i As Long, j As Long, k As Long, r As Long, count As Long, value As Double, id As Long
    Dim number As Long, source As String, message As String
    On Error GoTo Failed
    If Len(Dir$(path)) > 0 Then err.Raise 5, "RPX_Exchange", "OUTPUT_EXISTS"
    f = FreeFile: Open path For Binary Access Write As #f
    RPX_ExchangeMagic f, "P06VBA01": Put #f, , XN: Put #f, , XM: Put #f, , XB
    For i = 0 To XN - 1: value = XQ(i) * XQScale: Put #f, , value: Next i
    For i = 0 To XM - 1
        count = XEPtr(i + 1) - XEPtr(i): Put #f, , count: value = XERhs(i): Put #f, , value
        For j = XEPtr(i) To XEPtr(i + 1) - 1
            id = XEId(j): value = XEVal(j): Put #f, , id: Put #f, , value
        Next j
    Next i
    For k = 0 To XB - 1
        count = XC(k).dimn: Put #f, , count
        For r = 0 To XC(k).dimn - 1
            count = 0
            For j = 0 To XC(k).count - 1
                If XC(k).coef(r, j) <> 0# Then count = count + 1
            Next j
            Put #f, , count: value = XC(k).offset(r): Put #f, , value
            For j = 0 To XC(k).count - 1
                If XC(k).coef(r, j) <> 0# Then
                    id = XC(k).ids(j): value = XC(k).coef(r, j): Put #f, , id: Put #f, , value
                End If
            Next j
        Next r
    Next k
    Close #f: Exit Sub
Failed:
    number = err.number: source = err.source: message = err.description
    On Error Resume Next: Close #f: On Error GoTo 0
    err.Raise number, source, message
End Sub
Public Sub RPX_ExportExchangeContext(ByVal path As String)
    Dim fso As Object, s As Object, i As Long, j As Long, k As Long
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set s = fso.CreateTextFile(path, False, True)
    s.WriteLine "format=IEEE754_LE_HEX;policy=" & RPX_AnalysisPolicy() & ";load_phase=" & RPX_LoadPhase
    s.WriteLine "model" & vbTab & RPX_ExactModel()
    s.WriteLine "fs" & vbTab & RPX_ExactDouble(RStrengthFactor)
    s.WriteLine "q" & vbTab & RSubdivisions
    For i = 0 To XM - 1: s.WriteLine "row_norm" & vbTab & i & vbTab & RPX_ExactDouble(XEqRowNorm(i)): Next i
    For i = 0 To XN + XM - 1: s.WriteLine "permutation" & vbTab & i & vbTab & XPerm(i): Next i
    For i = 0 To rn - 1
        s.WriteLine "xy" & vbTab & i & vbTab & RPX_ExactDouble(RXY(i, 0)) & vbTab & RPX_ExactDouble(RXY(i, 1))
    Next i
    For i = 0 To re - 1
        s.WriteLine "tri" & vbTab & i & vbTab & RTri(i, 0) & vbTab & RTri(i, 1) & vbTab & RTri(i, 2) & vbTab & RMatId(i) & vbTab & RRigidId(i)
        For j = 0 To 2
            For k = 0 To 1: s.WriteLine "grad" & vbTab & i & vbTab & j & vbTab & k & vbTab & RPX_ExactDouble(RGrad(i, j, k)): Next k
        Next j
    Next i
    s.Close
End Sub
Public Sub RPX_ExportExchangeX(ByVal path As String)
    Dim f As Integer, i As Long, value As Double, number As Long, source As String, message As String
    On Error GoTo Failed
    If Len(Dir$(path)) > 0 Then err.Raise 5, "RPX_Exchange", "OUTPUT_EXISTS"
    f = FreeFile: Open path For Binary Access Write As #f
    RPX_ExchangeMagic f, "P06X0001": Put #f, , XN
    For i = 0 To XN - 1: value = xx(i): Put #f, , value: Next i
    Close #f: Exit Sub
Failed:
    number = err.number: source = err.source: message = err.description
    On Error Resume Next: Close #f: On Error GoTo 0
    err.Raise number, source, message
End Sub