Attribute VB_Name = "RPX_NativeBulk"
Option Explicit
' DLL_3: array-level kernels. No per-scalar calls from VBA to the DLL.
#If Win64 Then
Private Declare PtrSafe Function NScaling Lib "rpfem_native_x64.dll" Alias "rpx_scaling" (ByVal nb As Long, ByVal ns As Long, ByVal first As Long, ByVal last As Long, ByVal precise As Long, ByVal dims As LongPtr, ByVal starts As LongPtr, ByVal slack As LongPtr, ByVal dual As LongPtr, ByVal hh As LongPtr, ByVal wm As LongPtr, ByVal wi As LongPtr, ByVal si As LongPtr, ByVal v As LongPtr, ByVal flags As LongPtr, ByVal hd As LongPtr, ByRef preciseCount As Long) As Long
Private Declare PtrSafe Function NGProduct Lib "rpfem_native_x64.dll" Alias "rpx_gproduct" (ByVal nb As Long, ByVal ns As Long, ByVal nx As Long, ByVal first As Long, ByVal last As Long, ByVal dims As LongPtr, ByVal starts As LongPtr, ByVal ip As LongPtr, ByVal cp As LongPtr, ByVal ids As LongPtr, ByVal coef As LongPtr, ByVal offset As LongPtr, ByVal x As LongPtr, ByVal answer As LongPtr, ByVal withOffset As Long) As Long
Private Declare PtrSafe Function NGTAdd Lib "rpfem_native_x64.dll" Alias "rpx_gtadd" (ByVal nb As Long, ByVal ns As Long, ByVal nx As Long, ByVal first As Long, ByVal last As Long, ByVal dims As LongPtr, ByVal starts As LongPtr, ByVal ip As LongPtr, ByVal cp As LongPtr, ByVal ids As LongPtr, ByVal coef As LongPtr, ByVal x As LongPtr, ByVal answer As LongPtr, ByVal factor As Double) As Long
Private Declare PtrSafe Function NKKT Lib "rpfem_native_x64.dll" Alias "rpx_kkt_cones" (ByVal nb As Long, ByVal nnz As Long, ByVal first As Long, ByVal last As Long, ByVal dims As LongPtr, ByVal ip As LongPtr, ByVal cp As LongPtr, ByVal sp As LongPtr, ByVal slots As LongPtr, ByVal coef As LongPtr, ByVal hh As LongPtr, ByVal values As LongPtr) As Long
Private Declare PtrSafe Function NHOperator Lib "rpfem_native_x64.dll" Alias "rpx_hoperator" (ByVal n As Long, ByVal nx As Long, ByVal nb As Long, ByVal ns As Long, ByVal phase As Long, ByVal first As Long, ByVal last As Long, ByVal unshifted As Long, ByVal args As LongPtr) As Long
Private Declare PtrSafe Function NKInit Lib "rpfem_native_x64.dll" Alias "rpx_kkt_init" (ByVal count As Long, ByVal first As Long, ByVal last As Long, ByVal base As LongPtr, ByVal buffer As LongPtr) As Long
Private Declare PtrSafe Function NKDiag Lib "rpfem_native_x64.dll" Alias "rpx_kkt_diagonal" (ByVal n As Long, ByVal nx As Long, ByVal count As Long, ByVal first As Long, ByVal last As Long, ByVal diagonal As LongPtr, ByVal permutation As LongPtr, ByVal regularization As Double, ByVal buffer As LongPtr, ByVal scales As LongPtr) As Long
Private Declare PtrSafe Function NKScale Lib "rpfem_native_x64.dll" Alias "rpx_kkt_scale" (ByVal n As Long, ByVal count As Long, ByVal first As Long, ByVal last As Long, ByVal ap As LongPtr, ByVal ai As LongPtr, ByVal scales As LongPtr, ByVal buffer As LongPtr, ByVal original As LongPtr) As Long
Private Declare PtrSafe Function NCorrector Lib "rpfem_native_x64.dll" Alias "rpx_corrector" (ByVal nb As Long, ByVal ns As Long, ByVal phase As Long, ByVal first As Long, ByVal last As Long, ByVal dims As LongPtr, ByVal starts As LongPtr, ByVal target As Double, ByVal args As LongPtr) As Long
#End If
Private coneReady As Boolean, slotReady As Boolean
Private dims() As Long, starts() As Long, ip() As Long, cp() As Long, ids() As Long
Private coef() As Double, offsets() As Double, sp() As Long, slots() As Long
Private dummyFlags() As Boolean, dummyHD() As RPX_DD
Private packedN As Long, packedS As Long, packedB As Long
Private ntCalls As Long, gCalls As Long, gtCalls As Long, kktCalls As Long, opCalls As Long, correctionCalls As Long
Private guardReady As Boolean, guardClock As Double
Public Sub RPX_NativeBulkReset()
    coneReady = False: slotReady = False
    guardReady = False
End Sub
Public Sub RPX_NativeBulkResetCounts()
    ntCalls = 0: gCalls = 0: gtCalls = 0: kktCalls = 0: opCalls = 0
    correctionCalls = 0
End Sub
Public Function RPX_NativeBulkInfo() As String
    RPX_NativeBulkInfo = ";nt_bulk=" & ntCalls & ";g_bulk=" & gCalls & ";gt_bulk=" & gtCalls & ";kkt_bulk=" & kktCalls & ";hsd_operator=" & opCalls & ";corrector_bulk=" & correctionCalls
End Function
Public Function RPX_NativeCorrector(ByVal target As Double, ByRef si() As Double, ByRef v() As Double, ByRef ww() As Double, ByRef wi() As Double, ByRef ds() As Double, ByRef dz() As Double, ByRef rc() As Double) As Boolean
    Dim first As Long, last As Long, phase As Long, total As Long
    If Not RPX_NativeAvailable() Or XB <= 0 Or XS <= 0 Then Exit Function
#If Win64 Then
    Dim args(0 To 8) As LongPtr
    PrepareCones
    If LBound(si) <> 0 Or LBound(v) <> 0 Or LBound(ds) <> 0 Or LBound(dz) <> 0 Or LBound(rc) <> 0 Or LBound(XDual) <> 0 Or LBound(XSlack) <> 0 Then err.Raise 9, , "NATIVE_CORRECTOR_BASE"
    If UBound(si) < XS - 1 Or UBound(v) < XS - 1 Or UBound(ds) < XS - 1 Or UBound(dz) < XS - 1 Or UBound(rc) < XS - 1 Or UBound(XDual) < XS - 1 Or UBound(XSlack) < XS - 1 Then err.Raise 9, , "NATIVE_CORRECTOR_BUFFER"
    If LBound(ww, 1) <> 0 Or UBound(ww, 1) <> XB - 1 Or LBound(ww, 2) <> 0 Or UBound(ww, 2) <> 2 Or LBound(ww, 3) <> 0 Or UBound(ww, 3) <> 2 Then err.Raise 9, , "NATIVE_CORRECTOR_W_MATRIX"
    If LBound(wi, 1) <> 0 Or UBound(wi, 1) <> XB - 1 Or LBound(wi, 2) <> 0 Or UBound(wi, 2) <> 2 Or LBound(wi, 3) <> 0 Or UBound(wi, 3) <> 2 Then err.Raise 9, , "NATIVE_CORRECTOR_I_MATRIX"
    args(0) = VarPtr(XDual(0)): args(1) = VarPtr(si(0)): args(2) = VarPtr(v(0)): args(3) = VarPtr(ww(0, 0, 0)): args(4) = VarPtr(wi(0, 0, 0))
    args(5) = VarPtr(ds(0)): args(6) = VarPtr(dz(0)): args(7) = VarPtr(XSlack(0)): args(8) = VarPtr(rc(0))
    correctionCalls = correctionCalls + 1
    For phase = 0 To 1
        total = XS: If phase = 1 Then total = XB
        For first = 0 To total - 1 Step 128
            Call Pulse: last = first + 128: If last > total Then last = total
            RPX_NativeCheckCode NCorrector(XB, XS, phase, first, last, VarPtr(dims(0)), VarPtr(starts(0)), target, VarPtr(args(0)))
        Next first
    Next phase
    RPX_NativeCorrector = True
#End If
End Function
Public Function RPX_NativeHOperator(ByRef z() As RPX_DD, ByRef answer() As RPX_DD, ByRef scaleValues() As RPX_DDCoef, ByRef coefAt() As Long, ByRef coefficients() As RPX_DDCoef, ByRef blocks() As RPX_DDCoef, ByRef equalities() As RPX_DDCoef, ByRef positive As RPX_DDCoef, ByRef negative As RPX_DDCoef, ByVal unshifted As Boolean, ByRef opX() As RPX_DD, ByRef opV() As RPX_DD, ByRef opG() As RPX_DD, ByRef opH() As RPX_DD) As Boolean
    Dim n As Long, phase As Long, first As Long, last As Long, total As Long
    If Not RPX_NativeAvailable() Or XB <= 0 Or XS <= 0 Or XM <= 0 Then Exit Function
#If Win64 Then
    Dim args(0 To 19) As LongPtr
    PrepareCones
    n = XN + XM
    If LBound(z) <> 0 Or UBound(z) <> n - 1 Or UBound(answer) < n - 1 Or UBound(scaleValues) <> n - 1 Then err.Raise 9, "RPX_NativeBulk", "NATIVE_HSD_BUFFER"
    If UBound(opX) < n - 1 Or UBound(opV) < n - 1 Or UBound(opG) < XS - 1 Or UBound(opH) < XS - 1 Then err.Raise 9, "RPX_NativeBulk", "NATIVE_HSD_WORKSPACE"
    If LBound(answer) <> 0 Or LBound(scaleValues) <> 0 Or LBound(opX) <> 0 Or LBound(opV) <> 0 Or LBound(opG) <> 0 Or LBound(opH) <> 0 Then err.Raise 9, , "NATIVE_HSD_BASE"
    If LBound(XPerm) <> 0 Or UBound(XPerm) <> n - 1 Or LBound(coefAt) <> 0 Or UBound(coefAt) <> XB - 1 Or LBound(coefficients) <> 0 Then err.Raise 9, , "NATIVE_HSD_PATTERN_BUFFER"
    If LBound(blocks, 1) <> 0 Or UBound(blocks, 1) <> XB - 1 Or LBound(blocks, 2) <> 0 Or UBound(blocks, 2) <> 2 Or LBound(blocks, 3) <> 0 Or UBound(blocks, 3) <> 2 Then err.Raise 9, , "NATIVE_HSD_BLOCK_BUFFER"
    For phase = 0 To XB - 1
        If coefAt(phase) < 0 Or coefAt(phase) + dims(phase) * (ip(phase + 1) - ip(phase)) > UBound(coefficients) + 1 Then err.Raise 9, , "NATIVE_HSD_COEFFICIENT_BUFFER"
    Next phase
    If LBound(XEPtr) <> 0 Or UBound(XEPtr) <> XM Or XEPtr(0) <> 0 Or LBound(XEId) <> 0 Or LBound(equalities) <> 0 Then err.Raise 9, , "NATIVE_HSD_EQUALITY_BUFFER"
    If XEPtr(XM) < 0 Or XEPtr(XM) > UBound(XEId) + 1 Or XEPtr(XM) > UBound(equalities) + 1 Then err.Raise 9, , "NATIVE_HSD_EQUALITY_SIZE"
    args(0) = VarPtr(XPerm(0)): args(1) = VarPtr(scaleValues(0)): args(2) = VarPtr(dims(0)): args(3) = VarPtr(starts(0))
    args(4) = VarPtr(ip(0)): args(5) = VarPtr(ids(0)): args(6) = VarPtr(coefAt(0)): args(7) = VarPtr(coefficients(0))
    args(8) = VarPtr(blocks(0, 0, 0)): args(9) = VarPtr(XEPtr(0)): args(10) = VarPtr(XEId(0)): args(11) = VarPtr(equalities(0))
    args(12) = VarPtr(positive): args(13) = VarPtr(negative): args(14) = VarPtr(z(0)): args(15) = VarPtr(answer(0))
    args(16) = VarPtr(opX(0)): args(17) = VarPtr(opV(0)): args(18) = VarPtr(opG(0)): args(19) = VarPtr(opH(0))
    opCalls = opCalls + 1
    RPX_NativeCheckCode NHOperator(n, XN, XB, XS, -1, 0, 0, Abs(CLng(unshifted)), VarPtr(args(0)))
    For phase = 0 To 3
        total = n: If phase = 1 Then total = XB
        If phase = 2 Then total = XM
        For first = 0 To total - 1 Step 128
            Call Pulse: last = first + 128: If last > total Then last = total
            RPX_NativeCheckCode NHOperator(n, XN, XB, XS, phase, first, last, Abs(CLng(unshifted)), VarPtr(args(0)))
        Next first
    Next phase
    RPX_NativeHOperator = True
#End If
End Function
Private Sub Pulse()
    RPX_NativePulse
End Sub
Public Sub RPX_NativePulse()
    Dim nowClock As Double, elapsed As Double, interval As Double
    DoEvents
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    interval = RPX_StageIntervalValue(): nowClock = Timer
    If guardReady And interval > 0# Then
        elapsed = nowClock - guardClock: If elapsed < 0# Then elapsed = elapsed + 86400#
        If elapsed < interval Then Exit Sub
    End If
    guardReady = True: guardClock = nowClock
    RPX_AssertInputs "native_bulk"
End Sub
Private Sub PrepareCones(Optional ByVal needSlots As Boolean = False)
    Dim i As Long, j As Long, r As Long, k As Long, at As Long, idTotal As Long, coefTotal As Long, slotTotal As Long
    Dim pat As Long
    If coneReady Then
        If packedN <> XN Or packedS <> XS Or packedB <> XB Then err.Raise 5, "RPX_NativeBulk", "NATIVE_CONE_CONTEXT_MISMATCH"
    Else
        ReDim dims(0 To XB - 1): ReDim starts(0 To XB - 1): ReDim ip(0 To XB): ReDim cp(0 To XB)
        For i = 0 To XB - 1
            If i Mod 256 = 0 Then Pulse
            dims(i) = XC(i).dimn: starts(i) = XC(i).first
            If dims(i) <> 1 And dims(i) <> 3 Then err.Raise 5, "RPX_NativeBulk", "NATIVE_CONE_DIMENSION"
            If XC(i).count < 0 Or starts(i) < 0 Or starts(i) + dims(i) > XS Then err.Raise 9, "RPX_NativeBulk", "NATIVE_CONE_SIZE"
            idTotal = idTotal + XC(i).count: coefTotal = coefTotal + dims(i) * XC(i).count
            ip(i + 1) = idTotal: cp(i + 1) = coefTotal
        Next i
        RPX_CapacityCheck "native_cones", 12# * (XB + 1#) + 4# * idTotal + 8# * (coefTotal + XS)
        ReDim ids(0 To idTotal): ReDim coef(0 To coefTotal): ReDim offsets(0 To XS - 1)
        For i = 0 To XB - 1
            If i Mod 256 = 0 Then Pulse
            For j = 0 To XC(i).count - 1
                ids(ip(i) + j) = XC(i).ids(j)
                If ids(ip(i) + j) < 0 Or ids(ip(i) + j) >= XN Then err.Raise 9, "RPX_NativeBulk", "NATIVE_CONE_INDEX"
            Next j
            For r = 0 To dims(i) - 1
                offsets(starts(i) + r) = XC(i).offset(r)
                For j = 0 To XC(i).count - 1: coef(cp(i) + r * XC(i).count + j) = XC(i).coef(r, j): Next j
            Next r
        Next i
        packedN = XN: packedS = XS: packedB = XB: coneReady = True: slotReady = False
    End If
    If needSlots And Not slotReady Then
        ReDim sp(0 To XB)
        For i = 0 To XB - 1
            slotTotal = slotTotal + XC(i).count * (XC(i).count + 1) \ 2: sp(i + 1) = slotTotal
        Next i
        RPX_CapacityCheck "native_slots", 4# * (slotTotal + XB + 2#)
        ReDim slots(0 To slotTotal)
        For i = 0 To XB - 1
            If i Mod 256 = 0 Then Pulse
            pat = XC(i).Pattern
            If pat < 0 Or pat >= XB Then err.Raise 9, "RPX_NativeBulk", "NATIVE_SLOT_PATTERN"
            at = sp(i)
            For j = 0 To XC(i).count - 1
                For k = 0 To j
                    slots(at) = XC(pat).slots(j, k)
                    If slots(at) < 0 Or slots(at) > UBound(XKVal) Then err.Raise 9, "RPX_NativeBulk", "NATIVE_SLOT_INDEX"
                    at = at + 1
                Next k
            Next j
        Next i
        slotReady = True
    End If
End Sub
Public Function RPX_NativeScaling(ByVal precise As Boolean, ByRef hh() As Double, ByRef wm() As Double, ByRef wi() As Double, ByRef si() As Double, ByRef v() As Double, ByRef flags() As Boolean, ByRef hd() As RPX_DD, ByRef preciseCount As Long) As Boolean
    Dim first As Long, last As Long, code As Long
    If Not RPX_NativeAvailable() Or XB <= 0 Or XS <= 0 Then Exit Function
#If Win64 Then
    PrepareCones
    If LBound(XSlack) <> 0 Or UBound(XSlack) < XS - 1 Or LBound(XDual) <> 0 Or UBound(XDual) < XS - 1 Then err.Raise 9, "RPX_NativeBulk", "NATIVE_NT_BUFFER"
    If UBound(hh, 1) <> XB - 1 Or UBound(wm, 1) <> XB - 1 Or UBound(wi, 1) <> XB - 1 Or UBound(si) < XS - 1 Or UBound(v) < XS - 1 Then err.Raise 9, "RPX_NativeBulk", "NATIVE_NT_OUTPUT"
    If LBound(hh, 1) <> 0 Or LBound(wm, 1) <> 0 Or LBound(wi, 1) <> 0 Or LBound(si) <> 0 Or LBound(v) <> 0 Then err.Raise 9, , "NATIVE_NT_BASE"
    If LBound(hh, 2) <> 0 Or UBound(hh, 2) <> 2 Or LBound(hh, 3) <> 0 Or UBound(hh, 3) <> 2 Then err.Raise 9, , "NATIVE_NT_H_MATRIX"
    If LBound(wm, 2) <> 0 Or UBound(wm, 2) <> 2 Or LBound(wm, 3) <> 0 Or UBound(wm, 3) <> 2 Then err.Raise 9, , "NATIVE_NT_W_MATRIX"
    If LBound(wi, 2) <> 0 Or UBound(wi, 2) <> 2 Or LBound(wi, 3) <> 0 Or UBound(wi, 3) <> 2 Then err.Raise 9, , "NATIVE_NT_I_MATRIX"
    If precise Then
        If LBound(flags) <> 0 Or UBound(flags) <> XB - 1 Or LBound(hd, 1) <> 0 Or UBound(hd, 1) <> XB - 1 Then err.Raise 9, , "NATIVE_NT_DD_BUFFER"
        If LBound(hd, 2) <> 0 Or UBound(hd, 2) <> 2 Or LBound(hd, 3) <> 0 Or UBound(hd, 3) <> 2 Then err.Raise 9, , "NATIVE_NT_DD_MATRIX"
    End If
    If Not precise Then ReDim dummyFlags(0 To XB - 1): ReDim dummyHD(0 To XB - 1, 0 To 2, 0 To 2)
    ntCalls = ntCalls + 1
    For first = 0 To XB - 1 Step 64
        Call Pulse: last = first + 64: If last > XB Then last = XB
        If precise Then
            code = NScaling(XB, XS, first, last, 1, VarPtr(dims(0)), VarPtr(starts(0)), VarPtr(XSlack(0)), VarPtr(XDual(0)), VarPtr(hh(0, 0, 0)), VarPtr(wm(0, 0, 0)), VarPtr(wi(0, 0, 0)), VarPtr(si(0)), VarPtr(v(0)), VarPtr(flags(0)), VarPtr(hd(0, 0, 0)), preciseCount)
        Else
            code = NScaling(XB, XS, first, last, 0, VarPtr(dims(0)), VarPtr(starts(0)), VarPtr(XSlack(0)), VarPtr(XDual(0)), VarPtr(hh(0, 0, 0)), VarPtr(wm(0, 0, 0)), VarPtr(wi(0, 0, 0)), VarPtr(si(0)), VarPtr(v(0)), VarPtr(dummyFlags(0)), VarPtr(dummyHD(0, 0, 0)), preciseCount)
        End If
        RPX_NativeCheckCode code
    Next first
    RPX_NativeScaling = True
#End If
End Function
Public Function RPX_NativeGProduct(ByRef x() As Double, ByRef answer() As Double, ByVal withOffset As Boolean) As Boolean
    Dim first As Long, last As Long
    If Not RPX_NativeAvailable() Or XB <= 0 Or XS <= 0 Then Exit Function
#If Win64 Then
    PrepareCones
    If LBound(x) <> 0 Or UBound(x) < XN - 1 Or LBound(answer) <> 0 Or UBound(answer) < XS - 1 Then err.Raise 9, "RPX_NativeBulk", "NATIVE_G_BUFFER"
    gCalls = gCalls + 1
    For first = 0 To XB - 1 Step 256
        Call Pulse: last = first + 256: If last > XB Then last = XB
        RPX_NativeCheckCode NGProduct(XB, XS, XN, first, last, VarPtr(dims(0)), VarPtr(starts(0)), VarPtr(ip(0)), VarPtr(cp(0)), VarPtr(ids(0)), VarPtr(coef(0)), VarPtr(offsets(0)), VarPtr(x(0)), VarPtr(answer(0)), Abs(CLng(withOffset)))
    Next first
    RPX_NativeGProduct = True
#End If
End Function
Public Function RPX_NativeGTAdd(ByRef x() As Double, ByRef answer() As Double, ByVal factor As Double) As Boolean
    Dim first As Long, last As Long
    If Not RPX_NativeAvailable() Or XB <= 0 Or XS <= 0 Then Exit Function
#If Win64 Then
    PrepareCones
    If LBound(x) <> 0 Or UBound(x) < XS - 1 Or LBound(answer) <> 0 Or UBound(answer) < XN - 1 Then err.Raise 9, "RPX_NativeBulk", "NATIVE_GT_BUFFER"
    gtCalls = gtCalls + 1
    For first = 0 To XB - 1 Step 256
        Call Pulse: last = first + 256: If last > XB Then last = XB
        RPX_NativeCheckCode NGTAdd(XB, XS, XN, first, last, VarPtr(dims(0)), VarPtr(starts(0)), VarPtr(ip(0)), VarPtr(cp(0)), VarPtr(ids(0)), VarPtr(coef(0)), VarPtr(x(0)), VarPtr(answer(0)), factor)
    Next first
    RPX_NativeGTAdd = True
#End If
End Function
Public Function RPX_NativeKKTCones(ByRef hh() As Double) As Boolean
    Dim first As Long, last As Long, p As Long
    If Not RPX_NativeAvailable() Or XB <= 0 Then Exit Function
#If Win64 Then
    PrepareCones True
    If UBound(hh, 1) <> XB - 1 Or UBound(XKBase) <> UBound(XKVal) Then err.Raise 9, "RPX_NativeBulk", "NATIVE_KKT_BUFFER"
    If LBound(hh, 1) <> 0 Or LBound(hh, 2) <> 0 Or UBound(hh, 2) <> 2 Or LBound(hh, 3) <> 0 Or UBound(hh, 3) <> 2 Or LBound(XKBase) <> 0 Or LBound(XKVal) <> 0 Then err.Raise 9, , "NATIVE_KKT_BASE"
    For first = 0 To UBound(XKVal) Step 65536
        Call Pulse: last = first + 65536: If last > UBound(XKVal) + 1 Then last = UBound(XKVal) + 1
        RPX_NativeCheckCode NKInit(UBound(XKVal) + 1, first, last, VarPtr(XKBase(0)), VarPtr(XKVal(0)))
    Next first
    kktCalls = kktCalls + 1
    For first = 0 To XB - 1 Step 128
        Call Pulse: last = first + 128: If last > XB Then last = XB
        RPX_NativeCheckCode NKKT(XB, UBound(XKVal) + 1, first, last, VarPtr(dims(0)), VarPtr(ip(0)), VarPtr(cp(0)), VarPtr(sp(0)), VarPtr(slots(0)), VarPtr(coef(0)), VarPtr(hh(0, 0, 0)), VarPtr(XKVal(0)))
    Next first
    RPX_NativeKKTCones = True
#End If
End Function
Public Function RPX_NativeKKTTail(ByVal regularization As Double, ByRef original() As Double) As Boolean
    Dim n As Long, count As Long, first As Long, last As Long
    If Not RPX_NativeAvailable() Or XB <= 0 Then Exit Function
#If Win64 Then
    n = XN + XM: count = UBound(XKVal) + 1
    If LBound(XKPtr) <> 0 Or UBound(XKPtr) <> n Or XKPtr(n) <> count Or LBound(XKId) <> 0 Or UBound(XKId) <> count - 1 Then err.Raise 9, , "NATIVE_KKT_PATTERN_BUFFER"
    If LBound(XKScale) <> 0 Or UBound(XKScale) <> n - 1 Or LBound(XKDiag) <> 0 Or UBound(XKDiag) <> n - 1 Or LBound(XPerm) <> 0 Or UBound(XPerm) <> n - 1 Then err.Raise 9, , "NATIVE_KKT_VECTOR_BUFFER"
    If LBound(original) <> 0 Or UBound(original) <> count - 1 Then err.Raise 9, , "NATIVE_KKT_ORIGINAL_BUFFER"
    For first = 0 To n - 1 Step 256
        Call Pulse: last = first + 256: If last > n Then last = n
        RPX_NativeCheckCode NKDiag(n, XN, count, first, last, VarPtr(XKDiag(0)), VarPtr(XPerm(0)), regularization, VarPtr(XKVal(0)), VarPtr(XKScale(0)))
    Next first
    For first = 0 To n - 1 Step 256
        Call Pulse: last = first + 256: If last > n Then last = n
        RPX_NativeCheckCode NKScale(n, count, first, last, VarPtr(XKPtr(0)), VarPtr(XKId(0)), VarPtr(XKScale(0)), VarPtr(XKVal(0)), VarPtr(original(0)))
    Next first
    RPX_NativeKKTTail = True
#End If
End Function
