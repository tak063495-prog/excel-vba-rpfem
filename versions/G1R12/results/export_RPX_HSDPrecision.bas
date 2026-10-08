
Option Explicit
#Const HOPT12_INLINE_TWOSUM = 1
' HSD-only two-component arithmetic. No change to ordinary P4 arithmetic.
Public Type RPX_DD
    hi As Double
    lo As Double
End Type
Public Type RPX_DDCoef
    hi As Double
    lo As Double
    sh As Double
    sl As Double
End Type
Private Type HDoubleBits
    value As Double
End Type
Private Type HByteBits
    value(0 To 7) As Byte
End Type
Public Function HDD(ByVal value As Double) As RPX_DD
    HDD.hi = value
End Function
Private Sub TwoSum(ByVal a As Double, ByVal b As Double, ByRef s As Double, ByRef e As Double)
    Dim t As Double
    s = a + b: t = s - a: e = (a - (s - t)) + (b - t)
End Sub
Private Sub SplitDouble(ByVal a As Double, ByRef hi As Double, ByRef lo As Double)
    Dim bits As HDoubleBits, bytes As HByteBits
    ' Explicit binary truncation avoids the VBE expression evaluator retaining
    ' extra precision in splitter*(a) - (splitter*(a)-a).
    bits.value = a: LSet bytes = bits
    bytes.value(0) = 0: bytes.value(1) = 0: bytes.value(2) = 0
    bytes.value(3) = bytes.value(3) And 248
    LSet bits = bytes: hi = bits.value: lo = a - hi
End Sub
Public Sub HDSplit(ByVal value As Double, ByRef hiPart As Double, ByRef loPart As Double)
    SplitDouble value, hiPart, loPart
End Sub
Public Function HDAdd(ByRef a As RPX_DD, ByRef b As RPX_DD) As RPX_DD
    Dim s As Double, e As Double, t As Double, f As Double, u As Double, v As Double
    TwoSum a.hi, b.hi, s, e: TwoSum a.lo, b.lo, t, f
    TwoSum s, e + t, u, v: TwoSum u, v + f, HDAdd.hi, HDAdd.lo
End Function
Public Function HDNeg(ByRef a As RPX_DD) As RPX_DD
    HDNeg.hi = -a.hi: HDNeg.lo = -a.lo
End Function
Public Function HDSub(ByRef a As RPX_DD, ByRef b As RPX_DD) As RPX_DD
    Dim nb As RPX_DD
    nb = HDNeg(b): HDSub = HDAdd(a, nb)
End Function
Public Function HDMul(ByRef a As RPX_DD, ByRef b As RPX_DD) As RPX_DD
    Dim p As Double, e As Double, ah As Double, al As Double, bh As Double, bl As Double
    Dim t As Double
    p = a.hi * b.hi
    If Not RPX_SchurFinite(p) Then err.Raise 5, , "HSD_DD_NONFINITE_PRODUCT"
    If p = 0# And a.hi <> 0# And b.hi <> 0# Then err.Raise 5, , "HSD_DD_PRODUCT_UNDERFLOW"
    SplitDouble a.hi, ah, al: SplitDouble b.hi, bh, bl
    t = ah * bh: e = t - p
    t = ah * bl: e = e + t: t = al * bh: e = e + t: t = al * bl: e = e + t
    t = a.hi * b.lo: e = e + t: t = a.lo * b.hi: e = e + t
    t = a.lo * b.lo: e = e + t
    TwoSum p, e, HDMul.hi, HDMul.lo
End Function
Public Sub HDPrepare(ByVal value As Double, ByRef out As RPX_DDCoef)
    ' HDD(value) plus the SplitDouble of that hi. lo stays +0.
    out.hi = value: out.lo = 0#
    SplitDouble value, out.sh, out.sl
End Sub
Public Function HDMulCoef(ByRef a As RPX_DDCoef, ByRef b As RPX_DD) As RPX_DD
    ' HDMul when the left factor is HDD(value) and a.sh/a.sl are its split.
    Dim p As Double, e As Double, bh As Double, bl As Double, t As Double
    p = a.hi * b.hi
    If Not RPX_SchurFinite(p) Then err.Raise 5, , "HSD_DD_NONFINITE_PRODUCT"
    If p = 0# And a.hi <> 0# And b.hi <> 0# Then err.Raise 5, , "HSD_DD_PRODUCT_UNDERFLOW"
    SplitDouble b.hi, bh, bl
    t = a.sh * bh: e = t - p
    t = a.sh * bl: e = e + t: t = a.sl * bh: e = e + t: t = a.sl * bl: e = e + t
    t = a.hi * b.lo: e = e + t: t = a.lo * b.hi: e = e + t
    t = a.lo * b.lo: e = e + t
    TwoSum p, e, HDMulCoef.hi, HDMulCoef.lo
End Function
Public Function HDMulCoefKnown(ByRef a As RPX_DDCoef, ByRef b As RPX_DD, ByVal bh As Double, ByVal bl As Double) As RPX_DD
    ' Same product, checks, and error terms as HDMulCoef. bh/bl are SplitDouble of this b.hi.
    Dim p As Double, e As Double, t As Double
    p = a.hi * b.hi
    If Not RPX_SchurFinite(p) Then err.Raise 5, , "HSD_DD_NONFINITE_PRODUCT"
    If p = 0# And a.hi <> 0# And b.hi <> 0# Then err.Raise 5, , "HSD_DD_PRODUCT_UNDERFLOW"
    t = a.sh * bh: e = t - p
    t = a.sh * bl: e = e + t: t = a.sl * bh: e = e + t: t = a.sl * bl: e = e + t
    t = a.hi * b.lo: e = e + t: t = a.lo * b.hi: e = e + t
    t = a.lo * b.lo: e = e + t
    TwoSum p, e, HDMulCoefKnown.hi, HDMulCoefKnown.lo
End Function
Public Sub HDMulAddKnown(ByRef acc As RPX_DD, ByRef a As RPX_DDCoef, ByRef b As RPX_DD, ByVal bh As Double, ByVal bl As Double)
    ' HOPT12: same TwoSum statements, inlined.
    Dim p As Double, e As Double, t As Double
    Dim prodHi As Double, prodLo As Double
    Dim s As Double, f As Double, u As Double, v As Double
    Dim z As Double, arg3 As Double, arg4 As Double, outHi As Double, outLo As Double
    p = a.hi * b.hi
    If Not RPX_SchurFinite(p) Then err.Raise 5, , "HSD_DD_NONFINITE_PRODUCT"
    If p = 0# And a.hi <> 0# And b.hi <> 0# Then err.Raise 5, , "HSD_DD_PRODUCT_UNDERFLOW"
    t = a.sh * bh: e = t - p
    t = a.sh * bl: e = e + t: t = a.sl * bh: e = e + t: t = a.sl * bl: e = e + t
    t = a.hi * b.lo: e = e + t: t = a.lo * b.hi: e = e + t
    t = a.lo * b.lo: e = e + t
#If HOPT12_INLINE_TWOSUM Then
    z = p + e: arg3 = z - p
    prodHi = z: prodLo = (p - (z - arg3)) + (e - arg3)
    z = acc.hi + prodHi: arg3 = z - acc.hi
    s = z: e = (acc.hi - (z - arg3)) + (prodHi - arg3)
    z = acc.lo + prodLo: arg3 = z - acc.lo
    t = z: f = (acc.lo - (z - arg3)) + (prodLo - arg3)
    arg4 = e + t
    z = s + arg4: arg3 = z - s
    u = z: v = (s - (z - arg3)) + (arg4 - arg3)
    arg4 = v + f
    z = u + arg4: arg3 = z - u
    outHi = z: outLo = (u - (z - arg3)) + (arg4 - arg3)
    acc.hi = outHi: acc.lo = outLo
#Else
    TwoSum p, e, prodHi, prodLo
    TwoSum acc.hi, prodHi, s, e
    TwoSum acc.lo, prodLo, t, f
    TwoSum s, e + t, u, v
    TwoSum u, v + f, acc.hi, acc.lo
#End If
End Sub
Public Function HDMulCoefRight(ByRef a As RPX_DD, ByRef b As RPX_DDCoef) As RPX_DD
    ' HDMul when the right factor is HDD(value) and b.sh/b.sl are its split.
    Dim p As Double, e As Double, ah As Double, al As Double, t As Double
    p = a.hi * b.hi
    If Not RPX_SchurFinite(p) Then err.Raise 5, , "HSD_DD_NONFINITE_PRODUCT"
    If p = 0# And a.hi <> 0# And b.hi <> 0# Then err.Raise 5, , "HSD_DD_PRODUCT_UNDERFLOW"
    SplitDouble a.hi, ah, al
    t = ah * b.sh: e = t - p
    t = ah * b.sl: e = e + t: t = al * b.sh: e = e + t: t = al * b.sl: e = e + t
    t = a.hi * b.lo: e = e + t: t = a.lo * b.hi: e = e + t
    t = a.lo * b.lo: e = e + t
    TwoSum p, e, HDMulCoefRight.hi, HDMulCoefRight.lo
End Function
Public Function HDDiv(ByRef a As RPX_DD, ByRef b As RPX_DD) As RPX_DD
    Dim q As RPX_DD, r As RPX_DD, t As RPX_DD, k As Long
    If b.hi = 0# Then err.Raise 5, , "HSD_DD_ZERO_DIVISOR"
    q = HDD(a.hi / b.hi)
    For k = 1 To 2
        t = HDMul(b, q): r = HDSub(a, t): t = HDD((r.hi + r.lo) / b.hi): q = HDAdd(q, t)
    Next k
    HDDiv = q
End Function
Public Function HDValue(ByRef a As RPX_DD) As Double
    HDValue = a.hi + a.lo
    If Not RPX_SchurFinite(a.hi) Or Not RPX_SchurFinite(a.lo) Or Not RPX_SchurFinite(HDValue) Then err.Raise 5, , "HSD_DD_NONFINITE"
End Function
Public Sub HDProduct(ByRef ap() As Long, ByRef ai() As Long, ByRef av() As Double, ByRef x() As RPX_DD, ByRef result() As RPX_DD)
    Dim i As Long, j As Long, p As Long, a As RPX_DD, t As RPX_DD
    ReDim result(0 To UBound(x))
    For i = 0 To UBound(x)
        For p = ap(i) To ap(i + 1) - 1
            j = ai(p): a = HDD(av(p)): t = HDMul(a, x(i)): result(j) = HDAdd(result(j), t)
            If j <> i Then t = HDMul(a, x(j)): result(i) = HDAdd(result(i), t)
        Next p
    Next i
End Sub
Public Function RPX_TestHSDPrecision() As String
    Dim a As RPX_DD, b As RPX_DD, c As RPX_DD, d As RPX_DD, ah As Double, al As Double, bh As Double, bl As Double
    Dim cf As RPX_DDCoef, cr As RPX_DD
    On Error GoTo Failed
    a = HDD(1E+16): b = HDD(1#): c = HDAdd(a, b): c = HDSub(c, a)
    If c.hi <> 1# Or c.lo <> 0# Then err.Raise 513, , "DD_TWOSUM"
    a = HDD(1# + 2# ^ (-27)): b = HDD(1# - 2# ^ (-27)): c = HDMul(a, b)
    SplitDouble a.hi, ah, al: SplitDouble b.hi, bh, bl
    If c.hi <> 1# Or c.lo <> -(2# ^ (-54)) Then err.Raise 513, , "DD_TWOPROD hi=" & RPX_ExactDouble(c.hi) & ";lo=" & RPX_ExactDouble(c.lo) & ";ah=" & RPX_ExactDouble(ah) & ";al=" & RPX_ExactDouble(al) & ";bh=" & RPX_ExactDouble(bh) & ";bl=" & RPX_ExactDouble(bl)
    HDPrepare a.hi, cf: cr = HDMulCoef(cf, b)
    If cr.hi <> c.hi Or cr.lo <> c.lo Then err.Raise 513, , "DD_COEF_LEFT"
    HDSplit b.hi, bh, bl: cr = HDMulCoefKnown(cf, b, bh, bl)
    If cr.hi <> c.hi Or cr.lo <> c.lo Then err.Raise 513, , "DD_COEF_KNOWN"
    d = HDD(1#): cr = HDAdd(d, c): HDMulAddKnown d, cf, b, bh, bl
    If d.hi <> cr.hi Or d.lo <> cr.lo Then err.Raise 513, , "DD_MULADD"
    HDPrepare b.hi, cf: cr = HDMulCoefRight(c, cf): d = HDMul(c, b)
    If cr.hi <> d.hi Or cr.lo <> d.lo Then err.Raise 513, , "DD_COEF_RIGHT"
    HDPrepare 0.0000000001, cf: cr = HDMulCoef(cf, c): d = HDMul(HDD(0.0000000001), c)
    If cr.hi <> d.hi Or cr.lo <> d.lo Then err.Raise 513, , "DD_REG_COEF"
    d = HDDiv(c, a): c = HDSub(d, b)
    If Abs(HDValue(c)) > 1E-30 Then err.Raise 513, , "DD_DIVIDE"
    a = HDD(1E+290): b = HDD(1E-290): c = HDMul(a, b)
    If Abs(HDValue(c) - 1#) > 0.000000000000001 Then err.Raise 513, , "DD_SCALE"
    a = HDD(0#): c = HDMul(a, b)
    If HDValue(c) <> 0# Then err.Raise 513, , "DD_ZERO"
    RPX_TestHSDPrecision = "PASS TwoSum; TwoProd; cancellation; division; large scale; zero; coef cache; muladd"
    Exit Function
Failed:
    RPX_TestHSDPrecision = "FAIL " & err.description
End Function


Public Sub HDMulSubKnown(ByRef acc As RPX_DD, ByRef a As RPX_DD, ByRef b As RPX_DD, ByVal bh As Double, ByVal bl As Double)
    ' Same HDMul then HDSub arithmetic; bh/bl are the split of this b.hi.
    ' Split and rounding order are preserved. Only calls and UDT copies are removed.
    Dim p As Double, e As Double, ah As Double, al As Double, t As Double
    Dim prodHi As Double, prodLo As Double, s As Double, f As Double, u As Double, v As Double
    Dim z As Double, arg3 As Double, arg4 As Double, outHi As Double, outLo As Double
    p = a.hi * b.hi
    If Not RPX_SchurFinite(p) Then err.Raise 5, , "HSD_DD_NONFINITE_PRODUCT"
    If p = 0# And a.hi <> 0# And b.hi <> 0# Then err.Raise 5, , "HSD_DD_PRODUCT_UNDERFLOW"
    SplitDouble a.hi, ah, al
    t = ah * bh: e = t - p
    t = ah * bl: e = e + t: t = al * bh: e = e + t: t = al * bl: e = e + t
    t = a.hi * b.lo: e = e + t: t = a.lo * b.hi: e = e + t
    t = a.lo * b.lo: e = e + t
    z = p + e: arg3 = z - p
    prodHi = -z: prodLo = -((p - (z - arg3)) + (e - arg3))
    z = acc.hi + prodHi: arg3 = z - acc.hi
    s = z: e = (acc.hi - (z - arg3)) + (prodHi - arg3)
    z = acc.lo + prodLo: arg3 = z - acc.lo
    t = z: f = (acc.lo - (z - arg3)) + (prodLo - arg3)
    arg4 = e + t
    z = s + arg4: arg3 = z - s
    u = z: v = (s - (z - arg3)) + (arg4 - arg3)
    arg4 = v + f
    z = u + arg4: arg3 = z - u
    outHi = z: outLo = (u - (z - arg3)) + (arg4 - arg3)
    acc.hi = outHi: acc.lo = outLo
End Sub