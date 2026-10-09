Attribute VB_Name = "RPX_Solve"
Option Explicit
#Const HOPT13_DIR_SPLIT = 1
' version=20260925_2023 engine=20260925_warm_off
Private hh() As Double, ww() As Double, wi() As Double, sinvAll() As Double, vvAll() As Double
Private rp() As Double, rg() As Double, rd() As Double, rc() As Double
Private dx() As Double, dy() As Double, ds() As Double, dz() As Double
Private originalK() As Double
Private originalKReady As Boolean
Private originalKBound As Long
Private warmXX() As Double, warmXY() As Double, warmSlack() As Double, warmDual() As Double
Private warmValid As Boolean, warmKind As String
Private solveEngineAnnounced As Boolean
Private warmN As Long, warmM As Long, warmS As Long, warmB As Long
Public XStatus As String
Private hNTPrecise() As Boolean, hNTValues() As RPX_DD, hNTReady As Boolean
Private hTau As Double, hKappa As Double, dTau As Double, dKappa As Double
Private hOffset() As Double, hProduct() As Double, hRow() As Double, hColumnSolved() As Double
Private hDiagonal As Double, hScalarResidual As Double
Private hCertificateError As Double
Private hCaptureCount As Long
Private hPreciseFactor As Boolean, hPreciseAnswer() As RPX_DD
Private hBackend As String, hPivot As Boolean, hMatrixEpoch As Long, hFactorEpoch As Long, hArrowEpoch As Long, hRhsEpoch As Long
Private hFactorCounter As Long, hFactorMatrix As Long, hSnapshotKey As String
Private hInputEpoch As Long, hSolveEpoch As Long, hFailure As String, hColumn() As Double
Private hModelKey As String, hDenScale As Double, hDenRatio As Double
Private hMu As Double, hMua As Double, hSigma As Double, hTarget As Double, hChi As Double, hScalarRhs As Double
Private hQuality As RPX_HLinearQuality
Private hCaptureEnabled As Boolean, hCaptureMax As Long, hCaptureTotal As Long, hCaptureRun As String, hTrace As Long
Private hRole As String, hAttempt As Long, hDirectionPresent As Boolean, hConeDirectionPresent As Boolean
Private hBaseScaled() As Double, hBaseSolution() As Double, hBestSolution() As Double
Private hStepMethod As String, hSlackCone As Long, hDualCone As Long
Private hSlackHit As Double, hDualHit As Double, hLastAlpha As Double
Private hAnswerDD() As RPX_DD, hScaledDD() As RPX_DD, hArrowDD() As RPX_DD
Private hDirectionDD() As RPX_DD, hDirectionScaledDD() As RPX_DD
Private hDsDD() As RPX_DD, hDzDD() As RPX_DD, hDtDD As RPX_DD, hDkDD As RPX_DD
Private hBudget As Double
Private hSavedX() As Double, hSavedY() As Double, hSavedSlack() As Double, hSavedDual() As Double
Private hSavedTau As Double, hSavedKappa As Double
Private hLiveSaved As Boolean, hModelGuard As Boolean
Private Type HBitFloat
    bits As Double
End Type
Private Type HBitPair
    lo As Long
    hi As Long
End Type
Private refineCalls As Long
Private refineExit1 As Long
Private refineExit2 As Long
Private refineExit3 As Long
Private refineWorst As Double
Private hStructGen As Long
Private hCoefGen As Long
Private hEvalGen As Long
Private hScaleEpoch As Long
Private hHhEpoch As Long
Private hMatEpoch As Long
Private hSumEpoch As Long
Private hCoefReady As Boolean
Private hEvalReady As Boolean
Private hProjective As Boolean, hUnshifted As Boolean
Private hScaleReady As Boolean
Private hHhReady As Boolean
Private hMatReady As Boolean
Private hRegReady As Boolean
Private hRegPos As RPX_DDCoef
Private hRegNeg As RPX_DDCoef
Private hCoefDD() As RPX_DDCoef
Private hCoefAt() As Long
Private hEvalDD() As RPX_DDCoef
Private hScaleDD() As RPX_DDCoef
Private hHhDD() As RPX_DDCoef
Private hValDD() As RPX_DDCoef
Private hOrigDD() As RPX_DDCoef
Private hRowAbs() As Double
Private hOpX() As RPX_DD
Private hOpV() As RPX_DD
Private hOpGx() As RPX_DD
Private hOpHg() As RPX_DD
Private hOpXsh() As Double, hOpXsl() As Double
Private hOpGxsh() As Double, hOpGxsl() As Double
Private hBxSsh() As Double, hBxSsl() As Double
Private hBxOsh() As Double, hBxOsl() As Double
Private hBxRsh() As Double, hBxRsl() As Double
Private hBxOx() As RPX_DD
Private hBxRound() As RPX_DD
Private hBxP1() As RPX_DD
Private hBxP2() As RPX_DD
Private hBxP3() As RPX_DD
Private hBxDenom() As Double
Public Sub RPX_ResetPreparedState()
    ' Restoring compiled structure must not retain another model's numeric work state.
    originalKReady = False: warmValid = False
    hStructGen = hStructGen + 1: hCoefReady = False: hEvalReady = False
    Erase originalK: Erase warmXX: Erase warmXY: Erase warmSlack: Erase warmDual
    Erase xx: Erase xy: Erase XSlack: Erase XDual
    Erase dx: Erase dy: Erase ds: Erase dz
End Sub
Private Sub ResetRefineStats()
    refineCalls = 0
    refineExit1 = 0
    refineExit2 = 0
    refineExit3 = 0
    refineWorst = 0#
End Sub
Private Sub RecordRefine(ByVal passes As Long, ByVal maxError As Double)
    Dim used As Long
    used = passes
    If used < 1 Then used = 1
    If used > 3 Then used = 3
    refineCalls = refineCalls + 1
    If used = 1 Then
        refineExit1 = refineExit1 + 1
    ElseIf used = 2 Then
        refineExit2 = refineExit2 + 1
    Else
        refineExit3 = refineExit3 + 1
    End If
    If maxError > refineWorst Then refineWorst = maxError
End Sub
Private Sub EmitRefineSummary()
    RPX_DiagEvent "linear_refine_summary", "calls=" & refineCalls & ";exit1=" & refineExit1 & ";exit2=" & refineExit2 & ";exit3=" & refineExit3 & ";worst=" & refineWorst & ";threshold=1E-11;limit=3"
End Sub
Private Function Interior(ByRef a() As Double, ByRef da() As Double, ByVal stepSize As Double) As Boolean
    Dim i As Long, p As Long, t As Double, u As Double, v As Double, examined As Long
    For i = 0 To XB - 1
        examined = examined + 1
        p = XC(i).first: t = a(p) + stepSize * da(p)
        If t <= 0# Then
            RPX_DiagInteriorSample examined
            Exit Function
        End If
        If XC(i).dimn = 3 Then
            u = a(p + 1) + stepSize * da(p + 1): v = a(p + 2) + stepSize * da(p + 2)
            If t <= Sqr(u * u + v * v) Then
                RPX_DiagInteriorSample examined
                Exit Function
            End If
        End If
    Next i
    RPX_DiagInteriorSample examined
    Interior = True
End Function
Public Function ConeBoundaryHit(ByRef a() As Double, ByRef da() As Double, ByRef hit As Double, Optional ByRef limitingCone As Long = -1) As Boolean
    Dim i As Long, p As Long
    Dim s0 As Double, s1 As Double, s2 As Double
    Dim d0 As Double, d1 As Double, d2 As Double
    Dim aa As Double, bb As Double, cc As Double
    Dim disc As Double, rootQ As Double, root1 As Double, root2 As Double
    Dim cand As Double, t0 As Double, xscale As Double
    hit = 1E+300: limitingCone = -1
    For i = 0 To XB - 1
        p = XC(i).first
        s0 = a(p)
        d0 = da(p)
        If s0 <= 0# Then Exit Function
        If d0 < 0# Then
            cand = -s0 / d0
            If cand <= 0# Then Exit Function
            If cand < hit Then hit = cand: limitingCone = i
        End If
        If XC(i).dimn = 3 Then
            s1 = a(p + 1)
            s2 = a(p + 2)
            d1 = da(p + 1)
            d2 = da(p + 2)
            cc = s0 * s0 - s1 * s1 - s2 * s2
            If cc <= 0# Then Exit Function
            aa = d0 * d0 - d1 * d1 - d2 * d2
            bb = 2# * (s0 * d0 - s1 * d1 - s2 * d2)
            xscale = Abs(s0) + Abs(s1) + Abs(s2) + Abs(d0) + Abs(d1) + Abs(d2)
            If xscale < 1# Then xscale = 1#
            If Abs(aa) <= 1E-18 * xscale * xscale Then
                If Abs(bb) > 1E-18 * xscale * xscale Then
                    cand = -cc / bb
                    t0 = s0 + cand * d0
                    If cand > 0# And t0 >= 0# Then
                        If cand < hit Then hit = cand: limitingCone = i
                    End If
                End If
            Else
                disc = bb * bb - 4# * aa * cc
                If disc >= 0# Then
                    disc = Sqr(disc)
                    If bb >= 0# Then
                        rootQ = -0.5 * (bb + disc)
                    Else
                        rootQ = -0.5 * (bb - disc)
                    End If
                    If Abs(rootQ) <= 1E-300 Then Exit Function
                    root1 = rootQ / aa
                    root2 = cc / rootQ
                    t0 = s0 + root1 * d0
                    If root1 > 0# And t0 >= 0# Then
                        If root1 < hit Then hit = root1: limitingCone = i
                    End If
                    t0 = s0 + root2 * d0
                    If root2 > 0# And t0 >= 0# Then
                        If root2 < hit Then hit = root2: limitingCone = i
                    End If
                End If
            End If
        End If
    Next i
    ConeBoundaryHit = True
End Function
Private Function StepLimitByBisection() As Double
    Dim lo As Double, hi As Double, mid As Double, i As Long
    Dim slackOK As Boolean
    hi = 1#
    For i = 1 To 35
        mid = (lo + hi) / 2#
        slackOK = Interior(XSlack, ds, mid)
        If slackOK Then
            If Interior(XDual, dz, mid) Then
                lo = mid
            Else
                hi = mid
            End If
        Else
            hi = mid
        End If
    Next i
    StepLimitByBisection = lo
End Function
Private Function StepLimit() As Double
    Dim slackOK As Boolean, dualOK As Boolean
    Dim slackHit As Double, dualHit As Double
    Dim rawHit As Double, trialStep As Double
    hStepMethod = "": hSlackCone = -1: hDualCone = -1: hSlackHit = 0#: hDualHit = 0#
    RPX_DiagBegin dpStepLimit
    slackOK = Interior(XSlack, ds, 1#)
    If slackOK Then dualOK = Interior(XDual, dz, 1#)
    If slackOK And dualOK Then
        StepLimit = 1#: hStepMethod = "full"
        RPX_DiagStepLimitSample True, 0
        RPX_DiagEnd dpStepLimit
        Exit Function
    End If
    If ConeBoundaryHit(XSlack, ds, slackHit, hSlackCone) Then
        If ConeBoundaryHit(XDual, dz, dualHit, hDualCone) Then
            rawHit = slackHit
            If dualHit < rawHit Then rawHit = dualHit
            If rawHit > 0# And rawHit < 1E+300 Then
                trialStep = 0.99 * rawHit
                If trialStep > 1# Then trialStep = 1#
                slackOK = Interior(XSlack, ds, trialStep)
                If slackOK Then dualOK = Interior(XDual, dz, trialStep)
                If slackOK And dualOK Then
                    StepLimit = trialStep: hStepMethod = "boundary": hSlackHit = slackHit: hDualHit = dualHit
                    RPX_DiagStepLimitSample False, 0
                    RPX_DiagEnd dpStepLimit
                    Exit Function
                End If
            End If
        End If
    End If
    StepLimit = StepLimitByBisection(): hStepMethod = "bisection35": hSlackHit = slackHit: hDualHit = dualHit
    RPX_DiagStepLimitSample False, 35
    RPX_DiagEnd dpStepLimit
End Function
Private Function Residuals(Optional ByVal tau As Double = 1#) As Double
    Dim i As Long, p As Long, value As Double
    RPX_DiagBegin dpResiduals
    ReDim rp(0 To XM): ReDim rd(0 To XN - 1)
    RPX_GProduct xx, rg, True
    Dim cone As Long, component As Long
    For cone = 0 To XB - 1
        For component = 0 To XC(cone).dimn - 1
            p = XC(cone).first + component: rg(p) = rg(p) + (tau - 1#) * XC(cone).offset(component)
        Next component
    Next cone
    For i = 0 To XS - 1: rg(i) = rg(i) - XSlack(i): Next i
    For i = 0 To XN - 1: rd(i) = tau * XQ(i): Next i
    For i = 0 To XM - 1
        value = -tau * XERhs(i)
        For p = XEPtr(i) To XEPtr(i + 1) - 1
            value = value + XEVal(p) * xx(XEId(p))
            rd(XEId(p)) = rd(XEId(p)) + XEVal(p) * xy(i)
        Next p
        rp(i) = value
    Next i
    RPX_GTAdd XDual, rd, -1#
    XPrimalResidual = 0#: XDualResidual = 0#
    For i = 0 To XN - 1: If Abs(rd(i)) > XDualResidual Then XDualResidual = Abs(rd(i))
    Next i
    For i = 0 To XM - 1: If Abs(rp(i)) > XPrimalResidual Then XPrimalResidual = Abs(rp(i))
    Next i
    For i = 0 To XS - 1: If Abs(rg(i)) > XPrimalResidual Then XPrimalResidual = Abs(rg(i))
    Next i
    Residuals = XPrimalResidual: If XDualResidual > Residuals Then Residuals = XDualResidual
    RPX_DiagEnd dpResiduals
End Function
Private Sub Scaling(Optional ByVal preciseSOC As Boolean = False)
    Dim i As Long, p As Long, r As Long, j As Long
    Dim s(0 To 2) As Double, z(0 To 2) As Double, h(0 To 2, 0 To 2) As Double
    Dim w(0 To 2, 0 To 2) As Double, winv(0 To 2, 0 To 2) As Double, si(0 To 2) As Double, v(0 To 2) As Double
    Dim rs As Double, rz As Double, count As Long
    hNTReady = preciseSOC
    If preciseSOC Then
        ReDim hNTPrecise(0 To XB - 1): ReDim hNTValues(0 To XB - 1, 0 To 2, 0 To 2)
    End If
    For i = 0 To XB - 1
        p = XC(i).first
        If XC(i).dimn = 1 Then
            hh(i, 0, 0) = XDual(p) / XSlack(p): sinvAll(p) = 1# / XSlack(p)
        Else
            For r = 0 To 2: s(r) = XSlack(p + r): z(r) = XDual(p + r): Next r
            If preciseSOC Then
                rs = Sqr(s(1) * s(1) + s(2) * s(2)): rz = Sqr(z(1) * z(1) + z(2) * z(2))
                If s(0) <= rs Or z(0) <= rz Then err.Raise 5, "RPX_Solve", "NT_NONINTERIOR_POINT"
                hNTPrecise(i) = ((s(0) - rs) / s(0) < 0.000001 Or (z(0) - rz) / z(0) < 0.000001)
            End If
            If preciseSOC Then
                If hNTPrecise(i) Then
                    RPX_NTBlockPrecise s, z, h, w, winv, si, v
                    For r = 0 To 2
                        For j = 0 To 2: hNTValues(i, r, j) = RPX_NTPreciseH(r, j): Next j
                    Next r
                    count = count + 1
                Else
                    RPX_NTBlock s, z, h, w, winv, si, v
                End If
            Else
                RPX_NTBlock s, z, h, w, winv, si, v
            End If
            For r = 0 To 2
                sinvAll(p + r) = si(r): vvAll(p + r) = v(r)
                For j = 0 To 2
                    hh(i, r, j) = h(r, j): ww(i, r, j) = w(r, j): wi(i, r, j) = winv(r, j)
                Next j
            Next r
        End If
    Next i
End Sub
Private Sub FactorSystem(Optional ByVal regularization As Double = 0.000001, Optional ByVal assembleOnly As Boolean = False)
    Dim i As Long, j As Long, k As Long, p As Long, col As Long, nv As Long
    Dim value As Double, diag As Double, factorOK As Boolean
    Dim dimn As Long, cnt As Long, pat As Long
    Dim h00 As Double, h01 As Double, h02 As Double
    Dim h10 As Double, h11 As Double, h12 As Double
    Dim h20 As Double, h21 As Double, h22 As Double
    Dim c0 As Double, c1 As Double, c2 As Double
    Dim w0 As Double, w1 As Double, w2 As Double
    Dim k0 As Double, k1 As Double, k2 As Double
    RPX_DiagBegin dpKKT
    nv = XN + XM
    For p = 0 To UBound(XKVal): XKVal(p) = XKBase(p): Next p
    For i = 0 To XB - 1
        dimn = XC(i).dimn
        cnt = XC(i).count
        pat = XC(i).Pattern
        If dimn = 1 Then
            h00 = hh(i, 0, 0)
            For j = 0 To cnt - 1
                c0 = XC(i).coef(0, j)
                For k = 0 To j
                    p = XC(pat).slots(j, k)
                    XKVal(p) = XKVal(p) + c0 * h00 * XC(i).coef(0, k)
                Next k
            Next j
        Else
            h00 = hh(i, 0, 0): h01 = hh(i, 0, 1): h02 = hh(i, 0, 2)
            h10 = hh(i, 1, 0): h11 = hh(i, 1, 1): h12 = hh(i, 1, 2)
            h20 = hh(i, 2, 0): h21 = hh(i, 2, 1): h22 = hh(i, 2, 2)
            For j = 0 To cnt - 1
                c0 = XC(i).coef(0, j): c1 = XC(i).coef(1, j): c2 = XC(i).coef(2, j)
                w0 = h00 * c0 + h01 * c1 + h02 * c2
                w1 = h10 * c0 + h11 * c1 + h12 * c2
                w2 = h20 * c0 + h21 * c1 + h22 * c2
                For k = 0 To j
                    k0 = XC(i).coef(0, k): k1 = XC(i).coef(1, k): k2 = XC(i).coef(2, k)
                    value = k0 * w0 + k1 * w1 + k2 * w2
                    p = XC(pat).slots(j, k)
                    XKVal(p) = XKVal(p) + value
                Next k
            Next j
        End If
    Next i
    For i = 0 To nv - 1
        p = XKDiag(i)
        If XPerm(i) < XN Then XKVal(p) = XKVal(p) + regularization Else XKVal(p) = -regularization
        XKScale(i) = 1#
        diag = Abs(XKVal(p))
        If XPerm(i) < XN And diag > 1E-20 Then XKScale(i) = 1# / Sqr(diag)
    Next i
    If (Not originalKReady) Or (originalKBound <> UBound(XKVal)) Then
        ReDim originalK(0 To UBound(XKVal))
        originalKBound = UBound(XKVal)
        originalKReady = True
    End If
    For col = 0 To nv - 1
        For p = XKPtr(col) To XKPtr(col + 1) - 1
            originalK(p) = XKVal(p): XKVal(p) = XKVal(p) * XKScale(col) * XKScale(XKId(p))
        Next p
    Next col
    If assembleOnly Then RPX_DiagEnd dpKKT: Exit Sub
    RPX_DiagEnd dpKKT: RPX_DiagBegin dpLDL
    RPX_P4NewMatrix
    RPX_SchurFactor regularization
    If RPX_SchurAutoReady() Then
        factorOK = True
    Else
        factorOK = RPX_Numeric(XKPtr, XKId, XKVal)
        If factorOK Then P4Generic = P4Matrix
    End If
    RPX_DiagEnd dpLDL
    If Not factorOK Then
        Dim dumpFile As Integer
        If Len(XLogPath) > 0 Then
        dumpFile = FreeFile: Open XLogPath & ".matrix" For Output As #dumpFile
        Print #dumpFile, nv; ","; RPX_PivotFailure
        For col = 0 To nv - 1
            For p = XKPtr(col) To XKPtr(col + 1) - 1
                Print #dumpFile, XKId(p); ","; col; ","; XKVal(p)
            Next p
        Next col
        Close #dumpFile
        End If
        err.Raise 5, , "NEWTON_FACTORIZATION_FAILED pivot=" & RPX_PivotFailure
    End If
End Sub
Private Sub EnsureGeneric()
    Dim started As Double
    If P4Generic = P4Matrix Then Exit Sub
    started = RPX_DiagClock()
    RPX_TestHook "p4_generic_factor"
    If Not RPX_Numeric(XKPtr, XKId, XKVal) Then err.Raise 5, , "NEWTON_FACTORIZATION_FAILED"
    P4Generic = P4Matrix: P4Factors = P4Factors + 1
    RPX_P4Time 11, started
End Sub
Private Sub direction(ByVal kind As Long, ByRef info As RPX_DirectionInfo, Optional ByVal forceGeneric As Boolean = False)
    Dim i As Long, p As Long, r As Long, j As Long, nv As Long, it As Long
    Dim rhs() As Double, rhsP() As Double, answer() As Double, work() As Double, res() As Double, delta() As Double
    Dim value As Double, maxError As Double, backend As Long
    RPX_P4StartDirection kind
    RPX_DiagBegin dpLinearSolve
    nv = XN + XM: ReDim rhs(0 To nv - 1): ReDim work(0 To XS - 1)
    For i = 0 To XN - 1: rhs(i) = -rd(i): Next i
    For i = 0 To XM - 1: rhs(XN + i) = -rp(i): Next i
    For i = 0 To XB - 1
        p = XC(i).first
        For r = 0 To XC(i).dimn - 1
            work(p + r) = rc(p + r)
            For j = 0 To XC(i).dimn - 1: work(p + r) = work(p + r) + hh(i, r, j) * rg(p + j): Next j
        Next r
    Next i
    RPX_GTAdd work, rhs, -1#
    ReDim rhsP(0 To nv - 1)
    For i = 0 To nv - 1: rhsP(i) = rhs(XPerm(i)) * XKScale(i): Next i
    P4Current.attempted = 1
    If RPX_SchurAutoReady() And Not forceGeneric Then
        P4Current.attempted = 2
        If RPX_SchurAutoSolve(rhsP, answer, it, maxError) Then backend = 2: GoTo DirectionSolved
    End If
    EnsureGeneric
    backend = 1: P4Current.corrections = 0
    RPX_Backsolve rhsP, answer
    For it = 1 To 3
        RPX_SymmetricProduct XKPtr, XKId, XKVal, answer, res: maxError = 0#
        For i = 0 To nv - 1
            res(i) = rhsP(i) - res(i)
            If Abs(res(i)) > maxError Then maxError = Abs(res(i))
        Next i
        P4Current.checks = P4Current.checks + 1
        If maxError < 0.00000000001 Then Exit For
        RPX_Backsolve res, delta
        P4Current.corrections = P4Current.corrections + 1: P4Corrections = P4Corrections + 1
        For i = 0 To nv - 1: answer(i) = answer(i) + delta(i): Next i
    Next it
DirectionSolved:
    RPX_P4Direction backend, maxError
    info = P4Current
    RecordRefine P4Current.corrections, maxError
    RPX_SchurCompare rhsP, answer
    ReDim dx(0 To XN - 1): ReDim dy(0 To XM)
    For i = 0 To nv - 1
        If XPerm(i) < XN Then dx(XPerm(i)) = answer(i) * XKScale(i) Else dy(XPerm(i) - XN) = answer(i) * XKScale(i)
    Next i
    RPX_GProduct dx, ds: ReDim dz(0 To XS - 1)
    For i = 0 To XS - 1: ds(i) = ds(i) + rg(i): Next i
    For i = 0 To XB - 1
        p = XC(i).first
        For r = 0 To XC(i).dimn - 1
            value = -rc(p + r)
            For j = 0 To XC(i).dimn - 1: value = value - hh(i, r, j) * ds(p + j): Next j
            dz(p + r) = value
        Next r
    Next i
    RPX_DiagEnd dpLinearSolve
End Sub
Private Sub Corrector(ByVal target As Double)
    Dim i As Long, p As Long, r As Long, j As Long
    RPX_DiagBegin dpCorrector
    Dim a(0 To 2) As Double, b(0 To 2) As Double, w(0 To 2, 0 To 2) As Double, winv(0 To 2, 0 To 2) As Double
    Dim v(0 To 2) As Double, correction(0 To 2) As Double
    For i = 0 To XS - 1: rc(i) = XDual(i) - target * sinvAll(i): Next i
    For i = 0 To XB - 1
        p = XC(i).first
        If XC(i).dimn = 1 Then
            rc(p) = rc(p) + ds(p) * dz(p) / XSlack(p)
        Else
            For r = 0 To 2
                a(r) = ds(p + r): b(r) = dz(p + r): v(r) = vvAll(p + r)
                For j = 0 To 2: w(r, j) = ww(i, r, j): winv(r, j) = wi(i, r, j): Next j
            Next r
            RPX_NTCorrection v, w, winv, a, b, correction
            For r = 0 To 2: rc(p + r) = rc(p + r) + correction(r): Next r
        End If
    Next i
    RPX_DiagEnd dpCorrector
End Sub
Private Function CertificateReady(ByVal mode As String) As Boolean
    ' Gate rejection is a result; actual exceptions always propagate.
    If mode = "upper" Then Call RPX_AuditUpper(True)
    If mode = "lower" Then Call RPX_AuditLower(True)
    CertificateReady = True
    If mode = "upper" Or mode = "lower" Then CertificateReady = RPhysicalGateAccepted
    ' An ordinary candidate gate miss does not disguise an actual audit exception.
    RPX_AuditRejected = False
End Function
Private Sub StoreWarm(ByVal kind As String)
    Dim i As Long
    If XN <= 0 Or XS <= 0 Then Exit Sub
    ReDim warmXX(0 To XN - 1): ReDim warmXY(0 To XM)
    ReDim warmSlack(0 To XS - 1): ReDim warmDual(0 To XS - 1)
    For i = 0 To XN - 1: warmXX(i) = xx(i): Next i
    For i = 0 To XM - 1: warmXY(i) = xy(i): Next i
    For i = 0 To XS - 1
        warmSlack(i) = XSlack(i): warmDual(i) = XDual(i)
    Next i
    warmN = XN: warmM = XM: warmS = XS: warmB = XB
    warmKind = kind
    warmValid = True
End Sub
Private Function TryWarmStart(ByVal kind As String) As Boolean
    Dim i As Long, p As Long, t As Double, u As Double, v As Double, nrm As Double
    If Not warmValid Then Exit Function
    If warmKind <> kind Then Exit Function
    If warmN <> XN Or warmM <> XM Or warmS <> XS Or warmB <> XB Then Exit Function
    For i = 0 To XN - 1: xx(i) = warmXX(i): Next i
    For i = 0 To XM - 1: xy(i) = warmXY(i): Next i
    For i = 0 To XS - 1
        XSlack(i) = warmSlack(i): XDual(i) = warmDual(i)
    Next i
    For i = 0 To XB - 1
        p = XC(i).first
        If XC(i).dimn = 1 Then
            If XSlack(p) <= 0.0000000001 Then XSlack(p) = 1#
            If XDual(p) <= 0.0000000001 Then XDual(p) = 1#
        Else
            t = XSlack(p): u = XSlack(p + 1): v = XSlack(p + 2)
            nrm = Sqr(u * u + v * v)
            If t <= nrm + 0.0000000001 Then
                XSlack(p) = 1# + nrm
            End If
            t = XDual(p): u = XDual(p + 1): v = XDual(p + 2)
            nrm = Sqr(u * u + v * v)
            If t <= nrm + 0.0000000001 Then
                XDual(p) = 1# + nrm
            End If
        End If
    Next i
    TryWarmStart = True
End Function
Private Sub ColdPrimalDual()
    Dim i As Long, p As Long
    For i = 0 To XN - 1: xx(i) = 0#: Next i
    For i = 0 To XM - 1: xy(i) = 0#: Next i
    For i = 0 To XS - 1: XSlack(i) = 0#: XDual(i) = 0#: Next i
    For i = 0 To XB - 1
        p = XC(i).first: XSlack(p) = 1#: XDual(p) = 1#
    Next i
End Sub
Private Sub OptimizeCore(ByVal prepared As Boolean, ByVal auditMode As String)
    Dim i As Long, p As Long, mu As Double, mua As Double, alpha As Double, sigma As Double, valueScale As Double, f As Integer
    Dim usedWarm As Boolean
    Dim affine As RPX_DirectionInfo, correctorInfo As RPX_DirectionInfo, replayed As Boolean, replayStart As Double
    Call ResetRefineStats
    XStatus = "RUNNING"
    If Not prepared Then
        RPX_Compile
        originalKReady = False
        warmValid = False
    End If
    ReDim xx(0 To XN - 1): ReDim xy(0 To XM): ReDim XSlack(0 To XS - 1): ReDim XDual(0 To XS - 1)
    ReDim rc(0 To XS - 1): ReDim sinvAll(0 To XS - 1): ReDim vvAll(0 To XS - 1)
    ReDim hh(0 To XB - 1, 0 To 2, 0 To 2): ReDim ww(0 To XB - 1, 0 To 2, 0 To 2): ReDim wi(0 To XB - 1, 0 To 2, 0 To 2)
    usedWarm = False
    ' Nearby-F warm start is disabled. F=1 solution is not interior for F=2
    ' and caused CONIC_STEP_STALLED on the first prepared solve.
    Call ColdPrimalDual
    For XIteration = 0 To 149
        XResidual = Residuals(): mu = 0#: XObjective = 0#
        For i = 0 To XS - 1: mu = mu + XSlack(i) * XDual(i): Next i
        XGap = mu * XQScale: mu = mu / XB
        For i = 0 To XN - 1: XObjective = XObjective + XQ(i) * xx(i) * XQScale: Next i
        If usedWarm And XIteration = 0 Then
            If XPrimalResidual > 1000# Or XDualResidual > 1000# Or XGap > 1E+20 Then
                Call ColdPrimalDual
                usedWarm = False
                XResidual = Residuals(): mu = 0#: XObjective = 0#
                For i = 0 To XS - 1: mu = mu + XSlack(i) * XDual(i): Next i
                XGap = mu * XQScale: mu = mu / XB
                For i = 0 To XN - 1: XObjective = XObjective + XQ(i) * xx(i) * XQScale: Next i
            End If
        End If
        If Len(XLogPath) > 0 Then
            f = FreeFile: Open XLogPath For Append As #f
            Print #f, XIteration; ","; XObjective; ","; XResidual; ","; XGap; ","; RPX_FactorNNZ
            Close #f
        End If
        valueScale = Abs(XObjective): If valueScale < 1# Then valueScale = 1#
        If valueScale > 5# Then valueScale = 5#
        If XPrimalResidual < 0.00000001 And XDualResidual < 0.000001 And XGap < 0.000001 * valueScale Then
            If CertificateReady(auditMode) Then
                EmitRefineSummary
                StoreWarm auditMode
                Exit Sub
            End If
        End If
        Application.StatusBar = "RPFEM: conic iteration " & XIteration & " residual " & Format$(XResidual, "0.0E+00")
        DoEvents: If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        RPX_DiagEvent "iteration", "i=" & XIteration & ";primal=" & XPrimalResidual & ";dual=" & XDualResidual & ";gap=" & XGap
        RPX_DiagBegin dpScaling: Call Scaling: RPX_DiagEnd dpScaling: Call FactorSystem
        replayed = False
RetryDirections:
        RPX_SchurAutoFailed = False
        For i = 0 To XS - 1: rc(i) = XDual(i): Next i
        direction 1, affine, replayed: alpha = StepLimit(): mua = 0#
        For i = 0 To XS - 1: mua = mua + (XSlack(i) + alpha * ds(i)) * (XDual(i) + alpha * dz(i)): Next i
        sigma = mua / (XB * mu): If sigma > 1# Then sigma = 1#
        If sigma < 0# Then sigma = 0#
        Corrector mu * sigma ^ 3: direction 2, correctorInfo, (replayed Or affine.returned = 1)
        If (P4Consistent And affine.returned <> correctorInfo.returned) Or (Not P4Consistent And RPX_SchurAutoFailed) Then
            If replayed Then err.Raise 5, , "DIRECTION_REPLAY_LIMIT"
            replayed = True: replayStart = RPX_DiagClock(): P4Replays = P4Replays + 1
            If P4Trace > 0 Then RPX_DiagEvent "p4_pair_discard", "affine_rhs=" & affine.rhs & ";corrector_rhs=" & correctorInfo.rhs
            GoTo RetryDirections
        End If
        If replayed Then RPX_P4Time 10, replayStart
        alpha = 0.99 * StepLimit()
        If alpha < 0.000000000001 Then err.Raise 5, , "CONIC_STEP_STALLED"
        RPX_P4Commit affine, correctorInfo
        For i = 0 To XN - 1: xx(i) = xx(i) + alpha * dx(i): Next i
        For i = 0 To XM - 1: xy(i) = xy(i) + alpha * dy(i): Next i
        For i = 0 To XS - 1: XSlack(i) = XSlack(i) + alpha * ds(i): XDual(i) = XDual(i) + alpha * dz(i): Next i
    Next XIteration
    err.Raise 5, , "CONIC_NOT_CONVERGED"
End Sub
Private Sub HNativeCapture(ByRef rhs() As Double, ByRef scaled() As Double, ByRef solution() As Double, ByVal residual As Double, Optional ByVal force As Boolean = False)
    Dim f As Object, stream As Object, path As String, i As Long, j As Long, k As Long
    Dim eno As Long, esrc As String, edesc As String
    If Not hCaptureEnabled Then Exit Sub
    If hCaptureCount >= hCaptureMax Or hCaptureTotal >= hCaptureMax Then Exit Sub
    If Not force And residual <= 0.00000000001 Then
        If XIteration <> 0 Or hCaptureCount >= 3 Then Exit Sub
    End If
    hCaptureCount = hCaptureCount + 1: hCaptureTotal = hCaptureTotal + 1
    path = ThisWorkbook.path & "\hsd_linear_" & Format$(now, "yyyymmdd_hhnnss") & "_" & P4Solve & "_" & XIteration & "_" & hCaptureCount & ".tsv"
    On Error GoTo Failed
    Set f = CreateObject("Scripting.FileSystemObject")
    If f.FileExists(path) Or f.FileExists(path & ".partial") Then err.Raise 5, , "HSD_CAPTURE_COLLISION"
    Set stream = f.CreateTextFile(path & ".partial", False, False)
    stream.WriteLine "format" & vbTab & "IEEE754_LE_HEX_V1"
    stream.WriteLine "quality_operator" & vbTab & IIf(hUnshifted, "UNSHIFTED_GTHG_DD_SCALED_AND_ORIGINAL", "REGULARIZED_GTHG_DD_SCALED_AND_ORIGINAL")
    stream.WriteLine "context" & vbTab & P4Solve & vbTab & RPX_InputEpoch & vbTab & XIteration & vbTab & RPX_HSDAuditMode & vbTab & RPX_ExactDouble(RStrengthFactor)
    stream.WriteLine "generations" & vbTab & hMatrixEpoch & vbTab & hFactorEpoch & vbTab & hArrowEpoch & vbTab & hRhsEpoch & vbTab & hAttempt
    stream.WriteLine "role" & vbTab & hRole & vbTab & hPivot & vbTab & hDirectionPresent & vbTab & hQuality.valid
    stream.WriteLine "linear_quality" & vbTab & hQuality.checks & vbTab & hQuality.corrections & vbTab & RPX_ExactDouble(hQuality.bestResidual) & vbTab & RPX_ExactDouble(hQuality.finalResidual)
    stream.WriteLine "model_exact" & vbTab & hModelKey
    stream.WriteLine "load_phase" & vbTab & RPX_LoadPhase
    stream.WriteLine "regularization" & vbTab & RPX_ExactDouble(0.0000000001)
    stream.WriteLine "centering_present" & vbTab & (hRole = "corrector")
    stream.WriteLine "centering" & vbTab & RPX_ExactDouble(hMu) & vbTab & RPX_ExactDouble(hMua) & vbTab & RPX_ExactDouble(hSigma) & vbTab & RPX_ExactDouble(hTarget) & vbTab & RPX_ExactDouble(hChi)
    For i = 0 To hQuality.checks - 1: stream.WriteLine "refinement" & vbTab & i & vbTab & RPX_ExactDouble(hQuality.Residuals(i)): Next i
    If hRole <> "arrow" Then stream.WriteLine "scalar_rhs" & vbTab & RPX_ExactDouble(hScalarRhs)
    stream.WriteLine "factor_metrics_present" & vbTab & hPivot
    If hPivot Then stream.WriteLine "factor_metrics" & vbTab & RPX_ExactDouble(HLUBytes) & vbTab & RPX_ExactDouble(HLUGrowth) & vbTab & RPX_ExactDouble(HLUMinPivot) & vbTab & RPX_ExactDouble(HLUMaxL)
    stream.WriteLine "shape" & vbTab & XN & vbTab & XM & vbTab & XS & vbTab & XB
    stream.WriteLine "scalar" & vbTab & RPX_ExactDouble(hTau) & vbTab & RPX_ExactDouble(hKappa) & vbTab & RPX_ExactDouble(hScalarResidual) & vbTab & RPX_ExactDouble(residual)
    For i = 0 To XN + XM - 1
        stream.WriteLine "mapping" & vbTab & i & vbTab & XPerm(i) & vbTab & XInv(i) & vbTab & RPX_ExactDouble(XKScale(i))
        For j = XKPtr(i) To XKPtr(i + 1) - 1
            stream.WriteLine "matrix" & vbTab & j & vbTab & XKId(j) & vbTab & i & vbTab & RPX_ExactDouble(XKVal(j)) & vbTab & RPX_ExactDouble(originalK(j))
        Next j
    Next i
    HWriteVector stream, "best_solution", hBestSolution
    HWriteDD stream, "solution_dd", hScaledDD
    HWriteVector stream, "rhs", rhs: HWriteVector stream, "scaled_rhs", scaled: HWriteVector stream, "solution", solution
    HWriteVector stream, "x", xx: HWriteVector stream, "y", xy: HWriteVector stream, "slack", XSlack: HWriteVector stream, "dual", XDual
    HWriteVector stream, "rp", rp: HWriteVector stream, "rd", rd: HWriteVector stream, "rg", rg: HWriteVector stream, "offset", hOffset
    HWriteVector stream, "row", hRow: HWriteVector stream, "column", hColumn
    stream.WriteLine "arrow_present" & vbTab & (hArrowEpoch = hFactorEpoch And hFactorEpoch > 0)
    If hArrowEpoch = hFactorEpoch And hFactorEpoch > 0 Then HWriteVector stream, "arrow_solution", hColumnSolved
    If hArrowEpoch = hFactorEpoch And hFactorEpoch > 0 Then stream.WriteLine "border_diagonal" & vbTab & RPX_ExactDouble(hDiagonal)
    If hDirectionPresent Then
        HWriteDD stream, "direction_scaled_dd", hDirectionScaledDD
        stream.WriteLine "dtau_dd" & vbTab & RPX_ExactDouble(hDtDD.hi) & vbTab & RPX_ExactDouble(hDtDD.lo)
        HWriteVector stream, "dx", dx: HWriteVector stream, "dy", dy
        stream.WriteLine "cone_direction_present" & vbTab & hConeDirectionPresent
        If hConeDirectionPresent Then
            HWriteVector stream, "ds", ds: HWriteVector stream, "dz", dz
        End If
        stream.WriteLine "direction_scalar" & vbTab & RPX_ExactDouble(dTau) & vbTab & RPX_ExactDouble(dKappa)
    End If
    For i = 0 To XB - 1
        stream.WriteLine "cone" & vbTab & i & vbTab & XC(i).dimn & vbTab & XC(i).first
        For j = 0 To XC(i).dimn - 1
            For k = 0 To XC(i).count - 1
                stream.WriteLine "gcoef" & vbTab & XC(i).first + j & vbTab & XC(i).ids(k) & vbTab & RPX_ExactDouble(XC(i).coef(j, k))
            Next k
        Next j
        For j = 0 To 2
            For k = 0 To 2
                stream.WriteLine "nt" & vbTab & i & vbTab & j & vbTab & k & vbTab & RPX_ExactDouble(hh(i, j, k)) & vbTab & RPX_ExactDouble(ww(i, j, k)) & vbTab & RPX_ExactDouble(wi(i, j, k))
            Next k
        Next j
    Next i
    stream.Close: Set stream = Nothing
    If f.FileExists(path) Then err.Raise 5, "HNativeCapture", "HSD_CAPTURE_COLLISION"
    Name path & ".partial" As path
    Exit Sub
Failed:
    eno = err.number: esrc = err.source: edesc = err.description
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    On Error GoTo 0
    err.Raise eno, "HNativeCapture:" & esrc, "HSD_CAPTURE_FAILED: " & edesc & "; numerical_reason=" & hFailure
End Sub
Private Sub HWriteDD(ByVal stream As Object, ByVal name As String, ByRef values() As RPX_DD)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        stream.WriteLine name & vbTab & i & vbTab & RPX_ExactDouble(values(i).hi) & vbTab & RPX_ExactDouble(values(i).lo)
    Next i
End Sub
Private Sub HWriteVector(ByVal stream As Object, ByVal name As String, ByRef values() As Double)
    Dim i As Long
    For i = LBound(values) To UBound(values)
        stream.WriteLine name & vbTab & i & vbTab & RPX_ExactDouble(values(i))
    Next i
End Sub
Private Function HLiveKey() As String
    Dim parts() As String, p As Long, i As Long
    ReDim parts(0 To XN + XM + 2 * XS + 2)
    For i = 0 To XN - 1: parts(p) = RPX_ExactDouble(xx(i)): p = p + 1: Next i
    For i = 0 To XM: parts(p) = RPX_ExactDouble(xy(i)): p = p + 1: Next i
    For i = 0 To XS - 1: parts(p) = RPX_ExactDouble(XSlack(i)): p = p + 1: Next i
    For i = 0 To XS - 1: parts(p) = RPX_ExactDouble(XDual(i)): p = p + 1: Next i
    parts(p) = RPX_ExactDouble(hTau): parts(p + 1) = RPX_ExactDouble(hKappa)
    HLiveKey = Join(parts, "")
End Function
Private Function HLiveMatches() As Boolean
    Dim i As Long
    Dim liveBits As HBitFloat, savedBits As HBitFloat
    Dim livePair As HBitPair, savedPair As HBitPair
    If Not hLiveSaved Then Exit Function
    If LBound(xx) <> LBound(hSavedX) Or UBound(xx) <> UBound(hSavedX) Then Exit Function
    If LBound(xy) <> LBound(hSavedY) Or UBound(xy) <> UBound(hSavedY) Then Exit Function
    If LBound(XSlack) <> LBound(hSavedSlack) Or UBound(XSlack) <> UBound(hSavedSlack) Then Exit Function
    If LBound(XDual) <> LBound(hSavedDual) Or UBound(XDual) <> UBound(hSavedDual) Then Exit Function
    For i = 0 To XN - 1
        liveBits.bits = xx(i): savedBits.bits = hSavedX(i)
        LSet livePair = liveBits: LSet savedPair = savedBits
        If livePair.lo <> savedPair.lo Or livePair.hi <> savedPair.hi Then Exit Function
    Next i
    For i = 0 To XM
        liveBits.bits = xy(i): savedBits.bits = hSavedY(i)
        LSet livePair = liveBits: LSet savedPair = savedBits
        If livePair.lo <> savedPair.lo Or livePair.hi <> savedPair.hi Then Exit Function
    Next i
    For i = 0 To XS - 1
        liveBits.bits = XSlack(i): savedBits.bits = hSavedSlack(i)
        LSet livePair = liveBits: LSet savedPair = savedBits
        If livePair.lo <> savedPair.lo Or livePair.hi <> savedPair.hi Then Exit Function
    Next i
    For i = 0 To XS - 1
        liveBits.bits = XDual(i): savedBits.bits = hSavedDual(i)
        LSet livePair = liveBits: LSet savedPair = savedBits
        If livePair.lo <> savedPair.lo Or livePair.hi <> savedPair.hi Then Exit Function
    Next i
    liveBits.bits = hTau: savedBits.bits = hSavedTau
    LSet livePair = liveBits: LSet savedPair = savedBits
    If livePair.lo <> savedPair.lo Or livePair.hi <> savedPair.hi Then Exit Function
    liveBits.bits = hKappa: savedBits.bits = hSavedKappa
    LSet livePair = liveBits: LSet savedPair = savedBits
    If livePair.lo <> savedPair.lo Or livePair.hi <> savedPair.hi Then Exit Function
    HLiveMatches = True
End Function
Private Sub HContext()
    Dim mark As Long, failNumber As Long, failSource As String, failText As String
    Dim guardNow As Boolean
    mark = RPX_HoptMark()
    On Error GoTo HoptFailed
    ' Sheet identity and the model string are not arithmetic. HSD_STAGE_SECONDS
    ' gates them. A persistent edit is seen on the next due call. Sub-second
    ' edits that are put back are out of scope, same as hot progress.
    guardNow = RPX_ContextGuardDue()
    If guardNow Then
        RPX_HoptEnter HoptInputAssert
        RPX_AssertInputs "hsd_linear"
        RPX_HoptLeave HoptInputAssert
    End If
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    If hSolveEpoch <> P4Solve Or hInputEpoch <> RPX_InputEpoch Then err.Raise 5, , "HSD_CONTEXT_EPOCH_MISMATCH"
    If hFactorEpoch = 0 Or hFactorMatrix <> hMatrixEpoch Then err.Raise 5, , "HSD_FACTOR_CONTEXT_MISMATCH"
    RPX_HoptEnter HoptLiveKey
    If Not HLiveMatches() Then err.Raise 5, , "HSD_ITERATE_CONTEXT_MISMATCH"
    RPX_HoptLeave HoptLiveKey
    If guardNow Then
        RPX_HoptEnter HoptModelKey
        If hModelGuard Then
            If RPX_ExactModel() <> hModelKey Then err.Raise 5, , "HSD_MODEL_CONTEXT_MISMATCH"
        End If
        RPX_HoptLeave HoptModelKey
    End If
    Exit Sub
HoptFailed:
    failNumber = err.number: failSource = err.source: failText = err.description
    RPX_HoptUnwind mark
    err.Raise failNumber, failSource, failText
End Sub
Private Sub HFitDD(ByRef values() As RPX_DD, ByVal last As Long)
    Dim bad As Long, lb As Long, ub As Long
    If last < 0 Then Exit Sub
    bad = 0
    On Error Resume Next
    lb = LBound(values)
    If err.number <> 0 Then bad = 1: err.Clear
    ub = UBound(values)
    If err.number <> 0 Then bad = 1: err.Clear
    On Error GoTo 0
    If bad <> 0 Or lb <> 0 Or ub <> last Then ReDim values(0 To last)
End Sub
Private Sub HSplitDD(ByRef src() As RPX_DD, ByRef sh() As Double, ByRef sl() As Double, ByVal last As Long)
    Dim i As Long
    If last < 0 Then Exit Sub
    ReDim sh(0 To last)
    ReDim sl(0 To last)
    For i = 0 To last
        HDSplit src(i).hi, sh(i), sl(i)
    Next i
End Sub
Private Sub HClearDD(ByRef values() As RPX_DD, ByVal last As Long)
    Dim i As Long
    HFitDD values, last
    If last < 0 Then Exit Sub
    For i = 0 To last
        values(i).hi = 0#: values(i).lo = 0#
    Next i
End Sub
Private Sub HClearDouble(ByRef values() As Double, ByVal last As Long)
    Dim i As Long, bad As Long, lb As Long, ub As Long
    If last < 0 Then Exit Sub
    bad = 0
    On Error Resume Next
    lb = LBound(values)
    If err.number <> 0 Then bad = 1: err.Clear
    ub = UBound(values)
    If err.number <> 0 Then bad = 1: err.Clear
    On Error GoTo 0
    If bad <> 0 Or lb <> 0 Or ub <> last Then ReDim values(0 To last): Exit Sub
    For i = 0 To last
        values(i) = 0#
    Next i
End Sub
Private Sub HBuildScale(ByVal n As Long)
    Dim i As Long, bad As Long, ub As Long
    If n <= 0 Then Exit Sub
    If hScaleReady And hScaleEpoch = hMatrixEpoch Then
        bad = 0
        On Error Resume Next
        ub = UBound(hScaleDD)
        If err.number <> 0 Then bad = 1: err.Clear
        On Error GoTo 0
        If bad = 0 And ub = n - 1 Then Exit Sub
    End If
    ReDim hScaleDD(0 To n - 1)
    For i = 0 To n - 1
        HDPrepare XKScale(i), hScaleDD(i)
    Next i
    hScaleEpoch = hMatrixEpoch
    hScaleReady = True
End Sub
Private Sub HBuildHh()
    Dim k As Long, r As Long, j As Long, bad As Long, ub As Long
    If XB <= 0 Then
        Erase hHhDD: hHhEpoch = hMatrixEpoch: hHhReady = True
        Exit Sub
    End If
    If hHhReady And hHhEpoch = hMatrixEpoch Then
        bad = 0
        On Error Resume Next
        ub = UBound(hHhDD, 1)
        If err.number <> 0 Then bad = 1: err.Clear
        On Error GoTo 0
        If bad = 0 And ub = XB - 1 Then Exit Sub
    End If
    ReDim hHhDD(0 To XB - 1, 0 To 2, 0 To 2)
    For k = 0 To XB - 1
        For r = 0 To 2
            For j = 0 To 2
                HDPrepare hh(k, r, j), hHhDD(k, r, j)
                If hNTReady Then
                    If hNTPrecise(k) Then
                        hHhDD(k, r, j).hi = hNTValues(k, r, j).hi
                        hHhDD(k, r, j).lo = hNTValues(k, r, j).lo
                        HDSplit hNTValues(k, r, j).hi, hHhDD(k, r, j).sh, hHhDD(k, r, j).sl
                    End If
                End If
            Next j
        Next r
    Next k
    hHhEpoch = hMatrixEpoch
    hHhReady = True
End Sub
Private Sub HBuildCoef()
    Dim k As Long, r As Long, j As Long, at As Long, need As Long
    If XB <= 0 Then
        Erase hCoefDD: Erase hCoefAt
        hCoefReady = True: hCoefGen = hStructGen
        Exit Sub
    End If
    need = 0
    For k = 0 To XB - 1
        need = need + XC(k).dimn * XC(k).count
    Next k
    ReDim hCoefAt(0 To XB - 1)
    If need <= 0 Then Erase hCoefDD Else ReDim hCoefDD(0 To need - 1)
    at = 0
    For k = 0 To XB - 1
        hCoefAt(k) = at
        For r = 0 To XC(k).dimn - 1
            For j = 0 To XC(k).count - 1
                HDPrepare XC(k).coef(r, j), hCoefDD(at)
                at = at + 1
            Next j
        Next r
    Next k
    hCoefReady = True
    hCoefGen = hStructGen
End Sub
Private Sub HBuildEval()
    Dim p As Long, last As Long, bad As Long
    If hEvalReady And hEvalGen = hStructGen Then Exit Sub
    bad = 0
    On Error Resume Next
    last = UBound(XEVal)
    If err.number <> 0 Then bad = 1: err.Clear
    On Error GoTo 0
    If bad <> 0 Then
        Erase hEvalDD
    Else
        ReDim hEvalDD(0 To last)
        For p = 0 To last
            HDPrepare XEVal(p), hEvalDD(p)
        Next p
    End If
    hEvalReady = True
    hEvalGen = hStructGen
End Sub
Private Sub HEnsureDD(ByVal n As Long)
    If Not hRegReady Then
        HDPrepare 0.0000000001, hRegPos
        HDPrepare -0.0000000001, hRegNeg
        hRegReady = True
    End If
    HBuildScale n
    HBuildHh
    If (Not hCoefReady) Or hCoefGen <> hStructGen Then HBuildCoef
    If (Not hEvalReady) Or hEvalGen <> hStructGen Then HBuildEval
End Sub
Private Sub HBuildMatrix(ByVal n As Long)
    Dim i As Long, p As Long, r As Long, last As Long, bad As Long, ub As Long, rowUb As Long
    If n <= 0 Then Exit Sub
    If hMatReady And hMatEpoch = hMatrixEpoch And hSumEpoch = hMatrixEpoch Then
        bad = 0
        On Error Resume Next
        ub = UBound(hValDD)
        If err.number <> 0 Then bad = 1: err.Clear
        rowUb = UBound(hRowAbs)
        If err.number <> 0 Then bad = 1: err.Clear
        On Error GoTo 0
        If bad = 0 Then
            If ub = UBound(XKVal) And rowUb = n - 1 Then Exit Sub
        End If
    End If
    last = UBound(XKVal)
    ReDim hValDD(0 To last)
    ReDim hOrigDD(0 To last)
    For p = 0 To last
        HDPrepare XKVal(p), hValDD(p)
        HDPrepare originalK(p), hOrigDD(p)
    Next p
    ReDim hRowAbs(0 To n - 1)
    For i = 0 To n - 1
        For p = XKPtr(i) To XKPtr(i + 1) - 1
            r = XKId(p)
            hRowAbs(r) = hRowAbs(r) + Abs(XKVal(p))
            If r <> i Then hRowAbs(i) = hRowAbs(i) + Abs(XKVal(p))
        Next p
    Next i
    hMatEpoch = hMatrixEpoch
    hSumEpoch = hMatrixEpoch
    hMatReady = True
End Sub
Private Sub HBacksolve(ByRef rhs() As Double, ByRef answer() As Double)
    Dim mark As Long, failNumber As Long, failSource As String, failText As String
    mark = RPX_HoptMark()
    On Error GoTo HoptFailed
    RPX_DiagBegin dpLinearSolve
    RPX_HoptEnter HoptTriangular
    HContext
    If hPivot Then
        RPX_HLUSolve rhs, answer, hFactorEpoch
    ElseIf hPreciseFactor Then
        RPX_DDBacksolve rhs, hPreciseAnswer
        Dim i As Long
        ReDim answer(0 To XN + XM - 1)
        For i = 0 To XN + XM - 1: answer(i) = HDValue(hPreciseAnswer(i)): Next i
    Else
        RPX_Backsolve rhs, answer
    End If
    RPX_HoptLeave HoptTriangular
    RPX_DiagEnd dpLinearSolve
    Exit Sub
HoptFailed:
    failNumber = err.number: failSource = err.source: failText = err.description
    RPX_HoptUnwind mark
    RPX_DiagEnd dpLinearSolve
    err.Raise failNumber, failSource, failText
End Sub
Private Sub HOperator(ByRef z() As RPX_DD, ByRef result() As RPX_DD)
    ' Apply the original regularized equations before rounded G' H G assembly.
    ' The assembled/scaled KKT remains the preconditioner; delta is unchanged.
    ' Cached splits are HDD(value). Each output is still accumulated in the old order.
    Dim t As RPX_DD
    Dim n As Long, s As Long, i As Long, j As Long, k As Long, r As Long, p As Long, base As Long
    Dim hgSh As Double, hgSl As Double
    Dim mark As Long, failNumber As Long, failSource As String, failText As String
    mark = RPX_HoptMark()
    On Error GoTo HoptFailed
    RPX_HoptEnter HoptOperator
    n = XN + XM: s = XS
    HEnsureDD n
    HClearDD hOpX, n - 1
    HClearDD hOpV, n - 1
    HClearDD hOpGx, s - 1
    HClearDD hOpHg, s - 1
    HFitDD result, n - 1
    For i = 0 To n - 1: hOpX(XPerm(i)) = HDMulCoefRight(z(i), hScaleDD(i)): Next i
    HSplitDD hOpX, hOpXsh, hOpXsl, n - 1
    If s > 0 Then ReDim hOpGxsh(0 To s - 1): ReDim hOpGxsl(0 To s - 1)
    For k = 0 To XB - 1
        p = XC(k).first
        For r = 0 To XC(k).dimn - 1
            base = hCoefAt(k) + r * XC(k).count
            For j = 0 To XC(k).count - 1
                HDMulAddKnown hOpGx(p + r), hCoefDD(base + j), hOpX(XC(k).ids(j)), hOpXsh(XC(k).ids(j)), hOpXsl(XC(k).ids(j))
            Next j
        Next r
        For j = 0 To XC(k).dimn - 1
            HDSplit hOpGx(p + j).hi, hOpGxsh(p + j), hOpGxsl(p + j)
        Next j
        For r = 0 To XC(k).dimn - 1
            For j = 0 To XC(k).dimn - 1
                HDMulAddKnown hOpHg(p + r), hHhDD(k, r, j), hOpGx(p + j), hOpGxsh(p + j), hOpGxsl(p + j)
            Next j
            HDSplit hOpHg(p + r).hi, hgSh, hgSl
            base = hCoefAt(k) + r * XC(k).count
            For j = 0 To XC(k).count - 1
                i = XC(k).ids(j): HDMulAddKnown hOpV(i), hCoefDD(base + j), hOpHg(p + r), hgSh, hgSl
            Next j
        Next r
    Next k
    For i = 0 To XM - 1
        For p = XEPtr(i) To XEPtr(i + 1) - 1
            j = XEId(p): HDMulAddKnown hOpV(XN + i), hEvalDD(p), hOpX(j), hOpXsh(j), hOpXsl(j)
            HDMulAddKnown hOpV(j), hEvalDD(p), hOpX(XN + i), hOpXsh(XN + i), hOpXsl(XN + i)
        Next p
    Next i
    For i = 0 To n - 1
        j = XPerm(i)
        If Not hUnshifted Then
            If j < XN Then HDMulAddKnown hOpV(j), hRegPos, hOpX(j), hOpXsh(j), hOpXsl(j) Else HDMulAddKnown hOpV(j), hRegNeg, hOpX(j), hOpXsh(j), hOpXsl(j)
        End If
        result(i) = HDMulCoef(hScaleDD(i), hOpV(j))
    Next i
    RPX_HoptLeave HoptOperator
    Exit Sub
HoptFailed:
    failNumber = err.number: failSource = err.source: failText = err.description
    RPX_HoptUnwind mark
    err.Raise failNumber, failSource, failText
End Sub
Private Sub HBackward(ByRef b() As Double, ByRef x() As Double, ByRef quality As RPX_HLinearQuality, ByRef originalRhs() As Double)
    Dim i As Long, p As Long, r As Long, n As Long, anorm As Double, xnorm As Double, bnorm As Double, error As Double, ratio As Double
    Dim originalResidual As Double, roundedResidual As Double, worst As Long, worstValue As Double, t As RPX_DD
    Dim mark As Long, failNumber As Long, failSource As String, failText As String
    mark = RPX_HoptMark()
    On Error GoTo HoptFailed
    RPX_HoptEnter HoptBackward
    n = XN + XM
    HEnsureDD n
    HBuildMatrix n
    HFitDD hBxOx, n - 1
    HFitDD hBxRound, n - 1
    HClearDD hBxP1, n - 1
    HClearDD hBxP2, n - 1
    HClearDD hBxP3, n - 1
    HClearDouble hBxDenom, n - 1
    ' R4 B08: the three diagnostic products are not the acceptance test.
    ' trace=1 keeps the measured path. trace=0 skips them after matrix setup.
    If hTrace <= 0 Then
        RPX_HoptLeave HoptBackward
        Exit Sub
    End If
    For i = 0 To n - 1
        hBxOx(i) = HDMulCoefRight(hScaledDD(i), hScaleDD(i))
        hBxRound(i) = HDD(x(i))
    Next i
    ' One sparse pass. Each product still adds in row-then-entry order.
    ' All three products still run when hTrace is 0.
    RPX_HoptEnter HoptProduct
    HSplitDD hScaledDD, hBxSsh, hBxSsl, n - 1
    HSplitDD hBxOx, hBxOsh, hBxOsl, n - 1
    HSplitDD hBxRound, hBxRsh, hBxRsl, n - 1
    For i = 0 To n - 1
        For p = XKPtr(i) To XKPtr(i + 1) - 1
            r = XKId(p)
            HDMulAddKnown hBxP1(r), hValDD(p), hScaledDD(i), hBxSsh(i), hBxSsl(i)
            If r <> i Then HDMulAddKnown hBxP1(i), hValDD(p), hScaledDD(r), hBxSsh(r), hBxSsl(r)
            HDMulAddKnown hBxP2(r), hOrigDD(p), hBxOx(i), hBxOsh(i), hBxOsl(i)
            If r <> i Then HDMulAddKnown hBxP2(i), hOrigDD(p), hBxOx(r), hBxOsh(r), hBxOsl(r)
            HDMulAddKnown hBxP3(r), hValDD(p), hBxRound(i), hBxRsh(i), hBxRsl(i)
            If r <> i Then HDMulAddKnown hBxP3(i), hValDD(p), hBxRound(r), hBxRsh(r), hBxRsl(r)
            hBxDenom(r) = hBxDenom(r) + Abs(XKVal(p)) * Abs(x(i))
            If r <> i Then hBxDenom(i) = hBxDenom(i) + Abs(XKVal(p)) * Abs(x(r))
        Next p
    Next i
    RPX_HoptLeave HoptProduct
    For i = 0 To n - 1
        If hRowAbs(i) > anorm Then anorm = hRowAbs(i)
        If Abs(x(i)) > xnorm Then xnorm = Abs(x(i))
        If Abs(b(i)) > bnorm Then bnorm = Abs(b(i))
        ' Diagnostics use the actual stored b: rescaling-roundoff is displayed separately.
        t = HDSub(HDD(b(i)), hBxP1(i)): error = Abs(HDValue(t)): hBxDenom(i) = hBxDenom(i) + Abs(b(i))
        If error > worstValue Then worstValue = error: worst = i
        ratio = 0#
        If hBxDenom(i) > 0# Then ratio = error / hBxDenom(i) Else If error > 0# Then ratio = 1E+300
        If ratio > quality.componentBackward Then quality.componentBackward = ratio
        t = HDSub(HDD(originalRhs(XPerm(i))), hBxP2(i))
        If Abs(HDValue(t)) > originalResidual Then originalResidual = Abs(HDValue(t))
    Next i
    For i = 0 To n - 1
        t = HDSub(HDD(b(i)), hBxP3(i))
        If Abs(HDValue(t)) > roundedResidual Then roundedResidual = Abs(HDValue(t))
    Next i
    ratio = anorm * xnorm + bnorm
    If ratio > 0# Then quality.normwiseBackward = worstValue / ratio Else If worstValue > 0# Then quality.normwiseBackward = 1E+300
    If hTrace > 0 Then RPX_DiagEvent "hsd_multiscale_quality", "iteration=" & XIteration & ";scaled_stored_rhs=" & worstValue & ";originalK=" & originalResidual & ";rounded_scaled=" & roundedResidual & ";component=" & quality.componentBackward & ";normwise=" & quality.normwiseBackward & ";worst_original_id=" & XPerm(worst) & ";load_id=" & RLoadVariable & ";scale=" & XKScale(worst) & ";absolute_sum_rhs=" & hBxDenom(worst)
    RPX_HoptLeave HoptBackward
    Exit Sub
HoptFailed:
    failNumber = err.number: failSource = err.source: failText = err.description
    RPX_HoptUnwind mark
    err.Raise failNumber, failSource, failText
End Sub
Private Function HLinear(ByRef rhs() As Double, ByRef answer() As Double, ByRef quality As RPX_HLinearQuality) As Boolean
    Dim scaled() As Double, solution() As Double, residual() As Double, correction() As Double, best() As Double
    Dim xdd() As RPX_DD, bdd() As RPX_DD, prod() As RPX_DD, bestDD() As RPX_DD, t As RPX_DD
    Dim i As Long, j As Long, error As Double, previous As Double, emptyQuality As RPX_HLinearQuality
    Dim errorNumber As Long, errorSource As String, errorText As String
    On Error GoTo LinearFailed
    Erase answer: Erase hAnswerDD: Erase hScaledDD: quality = emptyQuality: hRhsEpoch = hRhsEpoch + 1
    quality.matrixEpoch = hMatrixEpoch: quality.factorEpoch = hFactorEpoch: quality.rhsEpoch = hRhsEpoch
    HContext
    ReDim scaled(0 To XN + XM - 1): ReDim bdd(0 To XN + XM - 1): ReDim xdd(0 To XN + XM - 1)
    ReDim residual(0 To XN + XM - 1)
    HEnsureDD XN + XM
    For i = 0 To XN + XM - 1
        If Not RPX_SchurFinite(rhs(XPerm(i))) Then err.Raise 5, , "HSD_NONFINITE_RHS"
        ' Retain the low part of S*b, including the arrow column.
        bdd(i) = HDMulCoefRight(HDD(rhs(XPerm(i))), hScaleDD(i)): scaled(i) = HDValue(bdd(i))
    Next i
    HBacksolve scaled, solution: quality.bestResidual = 1E+300
    If hPreciseFactor Then
        xdd = hPreciseAnswer
    Else
        For i = 0 To XN + XM - 1: xdd(i) = HDD(solution(i)): Next i
    End If
    For j = 0 To 8
        HOperator xdd, prod: error = 0#
        For i = 0 To XN + XM - 1
            t = HDSub(bdd(i), prod(i)): residual(i) = HDValue(t)
            If Abs(residual(i)) > error Then error = Abs(residual(i))
            If Abs(residual(i) / XKScale(i)) > error Then error = Abs(residual(i) / XKScale(i))
        Next i
        quality.Residuals(j) = error
        quality.checks = quality.checks + 1: quality.finalResidual = error
        If error < quality.bestResidual Then quality.bestResidual = error: bestDD = xdd
        If error <= 0.00000000001 Then Exit For
        If j > 0 And error > previous * 10# Then Exit For
        ' Extra checked LDL corrections are cheaper than refactoring. Continue
        ' only with >=4x contraction; pivot LU retains its three-correction cap.
        If j >= 3 Then
            If hPivot Or error > previous * 0.25 Then Exit For
        End If
        If j = 8 Then Exit For
        previous = error: HBacksolve residual, correction
        If hPreciseFactor Then
            For i = 0 To XN + XM - 1: xdd(i) = HDAdd(xdd(i), hPreciseAnswer(i)): Next i
        Else
            For i = 0 To XN + XM - 1: xdd(i) = HDAdd(xdd(i), HDD(correction(i))): Next i
        End If
        quality.corrections = quality.corrections + 1
    Next j
    If quality.bestResidual > 0.00000000001 And Not hPivot And Not hPreciseFactor Then
        Call HKRepair(bdd, bestDD, quality.bestResidual)
    End If
    ReDim best(0 To XN + XM - 1)
    For i = 0 To XN + XM - 1: best(i) = HDValue(bestDD(i)): Next i
    hBestSolution = best: hScaledDD = bestDD
    HBackward scaled, best, quality, rhs
    quality.valid = (quality.bestResidual <= 0.00000000001)
    If Not quality.valid Then hFailure = "HSD_LINEAR_QUALITY_FAILED"
    HNativeCapture rhs, scaled, best, error
    If hTrace > 0 Then RPX_DiagEvent "hsd_linear_quality", "iteration=" & XIteration & ";pivot=" & hPivot & ";matrix=" & hMatrixEpoch & ";rhs=" & hRhsEpoch & ";checks=" & quality.checks & ";corrections=" & quality.corrections & ";operator_unshifted=" & hUnshifted & ";operator_dd_max_scaled_original=" & quality.bestResidual & ";final=" & error & ";assembled_componentwise=" & quality.componentBackward
    If Not quality.valid Then Exit Function
    ReDim answer(0 To XN + XM - 1): ReDim hAnswerDD(0 To XN + XM - 1)
    For i = 0 To XN + XM - 1
        hAnswerDD(XPerm(i)) = HDMulCoefRight(bestDD(i), hScaleDD(i))
        answer(XPerm(i)) = HDValue(hAnswerDD(XPerm(i)))
    Next i
    hScaledDD = bestDD: hBaseScaled = scaled: hBaseSolution = best
    Call RPX_HoptEpoch(XIteration, hAttempt, hMatrixEpoch, hFactorEpoch, hRhsEpoch)
    HLinear = True
    Exit Function
LinearFailed:
    errorNumber = err.number: errorSource = err.source: errorText = err.description
    quality.valid = False: Erase answer: Erase hAnswerDD: Erase hScaledDD
    err.Raise errorNumber, errorSource, errorText
End Function
Private Function HFactor() As Boolean
    Dim ok As Boolean, bytes As Double
    Dim mark As Long, failNumber As Long, failSource As String, failText As String
    mark = RPX_HoptMark()
    On Error GoTo HoptFailed
    RPX_TestHook "hsd_factor"
    hFactorEpoch = 0: hArrowEpoch = 0
    If hPivot Then
        bytes = (UBound(XKVal) + 1) * 80# + (XN + XM) * 512# + XS * 320# + XB * 216#
        ok = RPX_HLUFactor(XKPtr, XKId, XKVal, hBudget, bytes, hFactorCounter + 1, False, True)
        If Not ok Then hFailure = "HSD_PIVOT_FACTORIZATION_FAILED"
    ElseIf hPreciseFactor Then
        RPX_DiagBegin dpLDL
        ok = HDDFactor()
        RPX_DiagEnd dpLDL
        If Not ok Then hFailure = "HSD_LINEAR_QUALITY_FAILED"
    Else
        RPX_DiagBegin dpLDL
        RPX_HoptEnter HoptLdl
        ok = RPX_Numeric(XKPtr, XKId, XKVal)
        RPX_HoptLeave HoptLdl
        RPX_DiagEnd dpLDL
        If Not ok Then hFailure = "HSD_LINEAR_QUALITY_FAILED"
    End If
    If ok Then hFactorCounter = hFactorCounter + 1: hFactorEpoch = hFactorCounter: hFactorMatrix = hMatrixEpoch
    Call RPX_HoptEpoch(XIteration, hAttempt, hMatrixEpoch, hFactorEpoch, hRhsEpoch)
    HFactor = ok
    Exit Function
HoptFailed:
    failNumber = err.number: failSource = err.source: failText = err.description
    RPX_HoptUnwind mark
    RPX_DiagEnd dpLDL
    err.Raise failNumber, failSource, failText
End Function
Private Sub HMultiply(ByRef a() As Double, ByRef value() As Double)
    Dim i As Long, p As Long, r As Long, j As Long
    ReDim value(0 To XS - 1)
    For i = 0 To XB - 1
        p = XC(i).first
        For r = 0 To XC(i).dimn - 1
            For j = 0 To XC(i).dimn - 1: value(p + r) = value(p + r) + hh(i, r, j) * a(p + j): Next j
        Next r
    Next i
End Sub
Private Function HArrow() As Boolean
    Dim column() As Double, i As Long
    ReDim column(0 To XN + XM - 1): ReDim hRow(0 To XN + XM - 1)
    HMultiply hOffset, hProduct
    RPX_GTAdd hProduct, column, 1#
    For i = 0 To XN - 1: hRow(i) = XQ(i) - column(i): column(i) = XQ(i) + column(i): Next i
    For i = 0 To XM - 1: column(XN + i) = -XERhs(i): hRow(XN + i) = XERhs(i): Next i
    hRole = "arrow": hColumn = column
    If Not HLinear(column, hColumnSolved, hQuality) Then Exit Function
    hArrowDD = hScaledDD
    hDiagonal = -hKappa / hTau
    For i = 0 To XS - 1: hDiagonal = hDiagonal - hOffset(i) * hProduct(i): Next i
    hArrowEpoch = hFactorEpoch: HArrow = True
End Function
Private Function HDirection(ByVal scalarCorrection As Double) As Boolean
    Dim rhs() As Double, work() As Double, answer() As Double, i As Long, scalarRhs As Double, denominator As Double
    hDirectionPresent = False: hConeDirectionPresent = False: hChi = scalarCorrection
    HContext
    HEnsureDD XN + XM
    If hArrowEpoch <> hFactorEpoch Then err.Raise 5, , "HSD_ARROW_CONTEXT_MISMATCH"
    ReDim rhs(0 To XN + XM - 1)
    HMultiply rg, work
    scalarRhs = -hScalarResidual + scalarCorrection
    For i = 0 To XS - 1
        work(i) = work(i) + rc(i): scalarRhs = scalarRhs + hOffset(i) * work(i)
    Next i
    For i = 0 To XN - 1: rhs(i) = -rd(i): Next i
    For i = 0 To XM - 1: rhs(XN + i) = -rp(i): Next i
    RPX_GTAdd work, rhs, -1#
    hScalarRhs = scalarRhs
    If Not HLinear(rhs, answer, hQuality) Then Exit Function
    Dim originalScalar As Double, denDD As RPX_DD, rhsDD As RPX_DD, rowDD As RPX_DD, t As RPX_DD
    Dim n As Long, k As Long, p As Long, r As Long, j As Long
    n = XN + XM: originalScalar = scalarRhs
    denDD = HDD(hDiagonal): rhsDD = HDD(scalarRhs): hDenScale = Abs(hDiagonal)
    For i = 0 To n - 1
        rowDD = HDMulCoefRight(HDD(hRow(XPerm(i))), hScaleDD(i))
        t = HDMul(rowDD, hArrowDD(i)): denDD = HDSub(denDD, t)
        hDenScale = hDenScale + Abs(HDValue(t))
        t = HDMul(rowDD, hScaledDD(i)): rhsDD = HDSub(rhsDD, t)
    Next i
    denominator = HDValue(denDD)
    hDenRatio = 0#: If hDenScale > 0# Then hDenRatio = Abs(denominator) / hDenScale
    If Abs(denominator) < 1E-290 Then hFailure = "HSD_BORDER_QUALITY_FAILED": Exit Function
    hDtDD = HDDiv(rhsDD, denDD)
    t = HDMul(HDD(hKappa / hTau), hDtDD): hDkDD = HDSub(HDD(-scalarCorrection), t)
    dTau = HDValue(hDtDD): dKappa = HDValue(hDkDD)
    ReDim dx(0 To XN - 1): ReDim dy(0 To XM)
    ReDim hDirectionDD(0 To n - 1): ReDim hDirectionScaledDD(0 To n - 1)
    For i = 0 To n - 1
        t = HDMul(hArrowDD(i), hDtDD): hDirectionScaledDD(i) = HDSub(hScaledDD(i), t)
        hDirectionDD(XPerm(i)) = HDMulCoefRight(hDirectionScaledDD(i), hScaleDD(i))
    Next i
    For i = 0 To XN - 1: dx(i) = HDValue(hDirectionDD(i)): Next i
    For i = 0 To XM - 1: dy(i) = HDValue(hDirectionDD(XN + i)): Next i
    If Not HBorder(rhs, originalScalar, scalarCorrection, denominator) Then
        If Not HRefineBorder(rhs, originalScalar, scalarCorrection, denDD) Then
            HNativeCapture rhs, hBaseScaled, hBaseSolution, hQuality.bestResidual, True
            Exit Function
        End If
    End If
    ' Preserve low components through G*dx, H*ds and the final state update.
    ' HOPT13: split hDirectionDD / cone hDsDD once. Do not reuse HOperator work arrays.
    ReDim hDsDD(0 To XS - 1): ReDim hDzDD(0 To XS - 1)
    ReDim ds(0 To XS - 1): ReDim dz(0 To XS - 1)
#If HOPT13_DIR_SPLIT Then
    Dim dirSh() As Double, dirSl() As Double
    Dim dsSh(0 To 2) As Double, dsSl(0 To 2) As Double
    Dim dimn As Long, id As Long
    ReDim dirSh(0 To n - 1): ReDim dirSl(0 To n - 1)
    For i = 0 To n - 1
        HDSplit hDirectionDD(i).hi, dirSh(i), dirSl(i)
    Next i
    For k = 0 To XB - 1
        p = XC(k).first
        dimn = XC(k).dimn
        For r = 0 To dimn - 1
            t = HDMul(HDD(hOffset(p + r)), hDtDD): hDsDD(p + r) = HDAdd(t, HDD(rg(p + r)))
            For j = 0 To XC(k).count - 1
                id = XC(k).ids(j)
                HDMulAddKnown hDsDD(p + r), hCoefDD(hCoefAt(k) + r * XC(k).count + j), hDirectionDD(id), dirSh(id), dirSl(id)
            Next j
        Next r
        For j = 0 To dimn - 1
            HDSplit hDsDD(p + j).hi, dsSh(j), dsSl(j)
        Next j
        For r = 0 To dimn - 1
            hDzDD(p + r) = HDD(-rc(p + r))
            For j = 0 To dimn - 1
                t = HDMulCoefKnown(hHhDD(k, r, j), hDsDD(p + j), dsSh(j), dsSl(j))
                hDzDD(p + r) = HDSub(hDzDD(p + r), t)
            Next j
            ds(p + r) = HDValue(hDsDD(p + r)): dz(p + r) = HDValue(hDzDD(p + r))
        Next r
    Next k
#Else
    For k = 0 To XB - 1
        p = XC(k).first
        For r = 0 To XC(k).dimn - 1
            t = HDMul(HDD(hOffset(p + r)), hDtDD): hDsDD(p + r) = HDAdd(t, HDD(rg(p + r)))
            For j = 0 To XC(k).count - 1
                t = HDMulCoef(hCoefDD(hCoefAt(k) + r * XC(k).count + j), hDirectionDD(XC(k).ids(j)))
                hDsDD(p + r) = HDAdd(hDsDD(p + r), t)
            Next j
        Next r
        For r = 0 To XC(k).dimn - 1
            hDzDD(p + r) = HDD(-rc(p + r))
            For j = 0 To XC(k).dimn - 1
                t = HDMulCoef(hHhDD(k, r, j), hDsDD(p + j)): hDzDD(p + r) = HDSub(hDzDD(p + r), t)
            Next j
            ds(p + r) = HDValue(hDsDD(p + r)): dz(p + r) = HDValue(hDzDD(p + r))
        Next r
    Next k
#End If
    hConeDirectionPresent = True
    HNativeCapture rhs, hBaseScaled, hBaseSolution, hQuality.bestResidual, True
    HDirection = True
End Function
Private Function HRefineBorder(ByRef rhs() As Double, ByVal scalarRhs As Double, ByVal chi As Double, ByRef denDD As RPX_DD) As Boolean
    ' G1R7: refine the coupled original operator, retaining the current K factor.
    ' HBorder still checks scaled/original residuals AND the kappa identity.
    Dim product() As RPX_DD, residual() As Double, correction() As Double
    Dim t As RPX_DD, scalar As RPX_DD, rowDD As RPX_DD, dt As RPX_DD
    Dim i As Long, pass As Long, n As Long, worst As Double, previous As Double
    n = XN + XM: previous = 1E+300
    ReDim residual(0 To n - 1)
    For pass = 1 To 3
        HContext
        HOperator hDirectionScaledDD, product
        scalar = HDSub(HDD(scalarRhs), HDMul(HDD(hDiagonal), hDtDD))
        worst = 0#
        For i = 0 To n - 1
            rowDD = HDMulCoefRight(HDD(hRow(XPerm(i))), hScaleDD(i))
            scalar = HDSub(scalar, HDMul(rowDD, hDirectionScaledDD(i)))
            t = HDSub(HDD(rhs(XPerm(i))), HDMul(HDD(hColumn(XPerm(i))), hDtDD))
            t = HDSub(HDMulCoef(hScaleDD(i), t), product(i))
            residual(i) = HDValue(t)
            If Abs(residual(i)) > worst Then worst = Abs(residual(i))
            If Abs(residual(i) / XKScale(i)) > worst Then worst = Abs(residual(i) / XKScale(i))
        Next i
        If Abs(HDValue(scalar)) > worst Then worst = Abs(HDValue(scalar))
        If pass > 1 And worst >= previous Then Exit Function
        previous = worst
        HBacksolve residual, correction
        dt = scalar
        For i = 0 To n - 1
            rowDD = HDMulCoefRight(HDD(hRow(XPerm(i))), hScaleDD(i))
            dt = HDSub(dt, HDMul(rowDD, HDD(correction(i))))
        Next i
        dt = HDDiv(dt, denDD)
        hDtDD = HDAdd(hDtDD, dt)
        For i = 0 To n - 1
            t = HDSub(HDD(correction(i)), HDMul(hArrowDD(i), dt))
            hDirectionScaledDD(i) = HDAdd(hDirectionScaledDD(i), t)
            hDirectionDD(XPerm(i)) = HDMulCoefRight(hDirectionScaledDD(i), hScaleDD(i))
        Next i
        hDkDD = HDSub(HDD(-chi), HDMul(HDD(hKappa / hTau), hDtDD))
        dTau = HDValue(hDtDD): dKappa = HDValue(hDkDD)
        For i = 0 To XN - 1: dx(i) = HDValue(hDirectionDD(i)): Next i
        For i = 0 To XM - 1: dy(i) = HDValue(hDirectionDD(XN + i)): Next i
        RPX_DiagEvent "hsd_border_refine", "iteration=" & XIteration & ";role=" & hRole & ";pass=" & pass & ";before=" & worst & ";factor=" & hFactorEpoch
        If HBorder(rhs, scalarRhs, chi, HDValue(denDD)) Then
            hFailure = "": HRefineBorder = True: Exit Function
        End If
    Next pass
End Function
Private Function HBorder(ByRef rhs() As Double, ByVal scalarRhs As Double, ByVal chi As Double, ByVal denominator As Double) As Boolean
    Dim product() As RPX_DD, t As RPX_DD, scalar As RPX_DD, rowDD As RPX_DD
    Dim i As Long, value As Double, worst As Double
    Dim mark As Long, failNumber As Long, failSource As String, failText As String
    mark = RPX_HoptMark()
    On Error GoTo HoptFailed
    RPX_HoptEnter HoptBorder
    HEnsureDD XN + XM
    HOperator hDirectionScaledDD, product
    scalar = HDSub(HDMul(HDD(hDiagonal), hDtDD), HDD(scalarRhs))
    For i = 0 To XN + XM - 1
        rowDD = HDMulCoefRight(HDD(hRow(XPerm(i))), hScaleDD(i))
        t = HDMul(rowDD, hDirectionScaledDD(i)): scalar = HDAdd(scalar, t)
        t = HDMul(HDD(hColumn(XPerm(i))), hDtDD): t = HDSub(t, HDD(rhs(XPerm(i))))
        t = HDMulCoef(hScaleDD(i), t): t = HDAdd(product(i), t)
        value = Abs(HDValue(t)): If value > worst Then worst = value
        value = Abs(HDValue(t) / XKScale(i)): If value > worst Then worst = value
    Next i
    value = Abs(HDValue(scalar)): If value > worst Then worst = value
    t = HDMul(HDD(hKappa / hTau), hDtDD): t = HDAdd(t, hDkDD): t = HDAdd(t, HDD(chi))
    value = Abs(HDValue(t)): If value > worst Then worst = value
    hDirectionPresent = True
    RPX_DiagEvent "hsd_border_quality", "iteration=" & XIteration & ";pivot=" & hPivot & ";dd_residual=" & worst & ";denominator=" & denominator & ";Dscale=" & hDenScale & ";Dratio=" & hDenRatio
    HBorder = (worst <= 0.00000000001)
    If Not HBorder Then hFailure = "HSD_BORDER_QUALITY_FAILED"
    RPX_HoptLeave HoptBorder
    Exit Function
HoptFailed:
    failNumber = err.number: failSource = err.source: failText = err.description
    RPX_HoptUnwind mark
    err.Raise failNumber, failSource, failText
End Function
Private Function HCommitInterior(ByRef ns() As Double, ByRef nz() As Double, ByVal nt As Double, ByVal nk As Double, ByRef cone As Long, ByRef side As String, ByRef margin As Double) As Boolean
    Dim i As Long, p As Long, radius As Double
    cone = -1
    If nt <= 0# Then side = "tau": margin = nt: Exit Function
    If nk <= 0# Then side = "kappa": margin = nk: Exit Function
    For i = 0 To XB - 1
        p = XC(i).first: radius = 0#
        If XC(i).dimn = 3 Then radius = Sqr(ns(p + 1) ^ 2 + ns(p + 2) ^ 2)
        margin = ns(p) - radius
        If margin <= 0# Then cone = i: side = "slack": Exit Function
        radius = 0#
        If XC(i).dimn = 3 Then radius = Sqr(nz(p + 1) ^ 2 + nz(p + 2) ^ 2)
        margin = nz(p) - radius
        If margin <= 0# Then cone = i: side = "dual": Exit Function
    Next i
    HCommitInterior = True
End Function
Private Sub HCommit(ByVal alpha As Double)
    Dim nx() As Double, ny() As Double, ns() As Double, nz() As Double, nt As Double, nk As Double, i As Long, p As Long
    Dim backtrack As Long, rejectedCone As Long, rejectedSide As String, margin As Double, originalAlpha As Double
    originalAlpha = alpha
    nx = xx: ny = xy: ns = XSlack: nz = XDual
    For backtrack = 0 To 24
    HContext
    For i = 0 To XS - 1: ns(i) = HDValue(HDAdd(HDD(XSlack(i)), HDMul(HDD(alpha), hDsDD(i)))): nz(i) = HDValue(HDAdd(HDD(XDual(i)), HDMul(HDD(alpha), hDzDD(i)))): Next i
    nt = HDValue(HDAdd(HDD(hTau), HDMul(HDD(alpha), hDtDD))): nk = HDValue(HDAdd(HDD(hKappa), HDMul(HDD(alpha), hDkDD)))
    For i = 0 To XS - 1
        If Not RPX_SchurFinite(ns(i)) Or Not RPX_SchurFinite(nz(i)) Then err.Raise 5, , "HSD_NONFINITE_COMMIT"
    Next i
    If Not RPX_SchurFinite(nt) Or Not RPX_SchurFinite(nk) Then err.Raise 5, , "HSD_NONFINITE_COMMIT"
    If HCommitInterior(ns, nz, nt, nk, rejectedCone, rejectedSide, margin) Then Exit For
    RPX_DiagEvent "hsd_commit_backtrack", "iteration=" & XIteration & ";trial=" & backtrack & ";alpha=" & alpha & ";side=" & rejectedSide & ";cone=" & rejectedCone & ";margin=" & margin & ";live_state_unchanged=True"
    alpha = alpha * 0.5
    If alpha < 0.000000000001 Then Exit For
    Next backtrack
    If Not HCommitInterior(ns, nz, nt, nk, rejectedCone, rejectedSide, margin) Then err.Raise 5, "RPX_Solve", "HSD_COMMIT_NOT_INTERIOR"
    hLastAlpha = alpha
    If backtrack > 0 Then RPX_DiagEvent "hsd_commit_recovered", "iteration=" & XIteration & ";backtracks=" & backtrack & ";original_alpha=" & originalAlpha & ";accepted_alpha=" & alpha
    For i = 0 To XN - 1: nx(i) = HDValue(HDAdd(HDD(xx(i)), HDMul(HDD(alpha), hDirectionDD(i)))): Next i
    For i = 0 To XM - 1: ny(i) = HDValue(HDAdd(HDD(xy(i)), HDMul(HDD(alpha), hDirectionDD(XN + i)))): Next i
    For i = 0 To XN - 1
        If Not RPX_SchurFinite(nx(i)) Then err.Raise 5, , "HSD_NONFINITE_COMMIT"
    Next i
    For i = 0 To XM - 1
        If Not RPX_SchurFinite(ny(i)) Then err.Raise 5, , "HSD_NONFINITE_COMMIT"
    Next i
    If hProjective Then
        Dim magnitude As Double, multiplier As Double
        magnitude = nt: If nk > magnitude Then magnitude = nk
        multiplier = 1#
        Do While magnitude > 4#: magnitude = magnitude / 2#: multiplier = multiplier / 2#: Loop
        If multiplier <> 1# Then
            For i = 0 To XN - 1: nx(i) = nx(i) * multiplier: Next i
            For i = 0 To XM - 1: ny(i) = ny(i) * multiplier: Next i
            For i = 0 To XS - 1: ns(i) = ns(i) * multiplier: nz(i) = nz(i) * multiplier: Next i
            nt = nt * multiplier: nk = nk * multiplier
            If nt <= 0# Or nk <= 0# Then err.Raise 5, "RPX_Solve", "HSD_PROJECTIVE_UNDERFLOW"
            RPX_DiagEvent "hsd_projective", "iteration=" & XIteration & ";multiplier=" & multiplier & ";all_homogeneous_variables=True;physical_point_unchanged=True"
        End If
    End If
    HContext
    xx = nx: xy = ny: XSlack = ns: XDual = nz: hTau = nt: hKappa = nk
End Sub
Private Function HStep() As Double
    HStep = StepLimit()
    Dim limiter As String
    limiter = "cone"
    If dTau < 0# Then
        If -hTau / dTau < HStep Then HStep = -hTau / dTau: limiter = "tau"
    End If
    If dKappa < 0# Then
        If -hKappa / dKappa < HStep Then HStep = -hKappa / dKappa: limiter = "kappa"
    End If
    hLastAlpha = HStep
    If hTrace > 0 Then RPX_DiagEvent "hsd_step", "iteration=" & XIteration & ";role=" & hRole & ";method=" & hStepMethod & ";limiter=" & limiter & ";slack_cone=" & hSlackCone & ";dual_cone=" & hDualCone & ";slack_hit=" & hSlackHit & ";dual_hit=" & hDualHit & ";alpha=" & HStep
End Function
Private Function ConeViolation(ByRef value() As Double) As Double
    Dim i As Long, p As Long, error As Double
    For i = 0 To XB - 1
        p = XC(i).first: error = -value(p)
        If XC(i).dimn = 3 Then error = error + Sqr(value(p + 1) ^ 2 + value(p + 2) ^ 2)
        If error > ConeViolation Then ConeViolation = error
    Next i
End Function
Private Function HCertificate() As Boolean
    Dim pc As Double, dc As Double, error As Double, value As Double, i As Long, p As Long, work() As Double
    hCertificateError = 1E+300
    For i = 0 To XN - 1: dc = dc + XQ(i) * xx(i): Next i
    For i = 0 To XM - 1: pc = pc + XERhs(i) * xy(i): Next i
    For i = 0 To XS - 1: pc = pc + hOffset(i) * XDual(i): Next i
    If pc < 0# Then
        error = ConeViolation(XDual)
        For i = 0 To XN - 1
            value = Abs(rd(i) - XQ(i) * hTau): If value > error Then error = value
        Next i
        hCertificateError = error / (-pc)
        If hCertificateError < 0.00000001 Then
            If TryHsdCandidate("PRIMAL_INFEASIBILITY_CERT", -pc, pc, dc, hCertificateError, XObjective, XGap) Then HCertificate = True: Exit Function
        End If
    End If
    If dc < 0# Then
        RPX_GProduct xx, work: error = ConeViolation(work)
        For i = 0 To XM - 1
            value = Abs(rp(i) + XERhs(i) * hTau): If value > error Then error = value
        Next i
        If error / (-dc) < hCertificateError Then hCertificateError = error / (-dc)
        If error / (-dc) < 0.00000001 Then
            HCertificate = TryHsdCandidate("LOAD_RECESSION_CERT", -dc, pc, dc, error / (-dc), XObjective, XGap)
        End If
    End If
End Function
Private Function TryHsdCandidate(ByVal kind As String, ByVal divisor As Double, ByVal pc As Double, ByVal dc As Double, ByVal certError As Double, ByVal objective As Double, ByVal gap As Double) As Boolean
    Dim candidate As RPX_HsdCandidateState
    RPX_HSDObserve kind, hTau, hKappa, divisor, pc, dc, certError
    If Not RPX_HSDBuildCandidate(candidate, kind, divisor, hTau, hKappa, objective, gap) Then Exit Function
    candidate.pc = pc: candidate.dc = dc: candidate.certificateError = certError
    If Not RPX_HSDValidateCandidate(candidate) Then Exit Function
    RPX_HSDCommitCandidate candidate
    TryHsdCandidate = True
End Function
Private Sub OptimizeHSDCore(ByVal prepared As Boolean, ByVal auditMode As String)
    Dim i As Long, p As Long, j As Long, mu As Double, gap As Double, objective As Double, alpha As Double, mua As Double, sigma As Double, target As Double, correction As Double
    Dim solveStarted As Double, solveBudget As Double
    solveBudget = CDbl(RPX_Setting("HSD_SOLVE_SECONDS", 120))
    If solveBudget <= 0# Or Not RPX_SchurFinite(solveBudget) Then err.Raise 5, "RPX_Solve", "HSD_SOLVE_BUDGET_INVALID"
    Dim representation As Double, operatorSetting As Double
    representation = CDbl(RPX_Setting("HSD_PROJECTIVE_NORMALIZATION", 1))
    operatorSetting = CDbl(RPX_Setting("HSD_ORIGINAL_OPERATOR", 0))
    If (representation <> 0# And representation <> 1#) Or (operatorSetting <> 0# And operatorSetting <> 1#) Then err.Raise 5, "RPX_Solve", "HSD_REPRESENTATION_INVALID"
    hProjective = (representation = 1#): hUnshifted = (operatorSetting = 1#)
    solveStarted = RPX_DiagClock()
    Call ResetRefineStats
    XStatus = "RUNNING": XObjective = 0#: If Not prepared Then RPX_Compile: originalKReady = False: warmValid = False
    hStructGen = hStructGen + 1
    ReDim xx(0 To XN - 1): ReDim xy(0 To XM): ReDim XSlack(0 To XS - 1): ReDim XDual(0 To XS - 1)
    ReDim rc(0 To XS - 1): ReDim sinvAll(0 To XS - 1): ReDim vvAll(0 To XS - 1): ReDim hOffset(0 To XS - 1)
    ReDim hh(0 To XB - 1, 0 To 2, 0 To 2): ReDim ww(0 To XB - 1, 0 To 2, 0 To 2): ReDim wi(0 To XB - 1, 0 To 2, 0 To 2)
    Call ColdPrimalDual
    For i = 0 To XB - 1
        p = XC(i).first
        For j = 0 To XC(i).dimn - 1: hOffset(p + j) = XC(i).offset(j): Next j
    Next i
    hCaptureCount = 0
    Dim captureSetting As Double, maxSetting As Double, traceSetting As Double
    captureSetting = CDbl(RPX_Setting("HSD_LINEAR_CAPTURE", 0))
    maxSetting = CDbl(RPX_Setting("HSD_LINEAR_CAPTURE_MAX", 8))
    traceSetting = CDbl(RPX_Setting("HSD_LINEAR_TRACE", 1))
    If captureSetting <> 0# And captureSetting <> 1# Then err.Raise 5, , "HSD_CONFIGURATION_INVALID"
    If maxSetting <> Fix(maxSetting) Or traceSetting <> Fix(traceSetting) Then err.Raise 5, , "HSD_CONFIGURATION_INVALID"
    hCaptureEnabled = (captureSetting = 1#)
    hCaptureMax = CLng(maxSetting): hTrace = CLng(traceSetting)
    If hCaptureMax < 0 Or hTrace < 0 Or hTrace > 2 Then err.Raise 5, , "HSD_CONFIGURATION_INVALID"
    If hCaptureRun <> RPX_DiagPath Then hCaptureTotal = 0: hCaptureRun = RPX_DiagPath
    hBackend = UCase$(CStr(RPX_Setting("HSD_LINEAR_BACKEND", "AUTO_PIVOT")))
    If hBackend <> "LDL_CHECKED" And hBackend <> "PIVOT_ONLY" And hBackend <> "AUTO_PIVOT" Then err.Raise 5, , "HSD_BACKEND_INVALID"
    hBudget = CDbl(RPX_Setting("HSD_STABLE_MAX_BYTES", 268435456))
    If hBudget <= 0# Then err.Raise 5, , "HSD_BUDGET_INVALID"
    hInputEpoch = RPX_InputEpoch: hSolveEpoch = P4Solve
    hModelKey = "": hModelGuard = False
    If Len(auditMode) > 0 Then hModelKey = RPX_ExactModel(): hModelGuard = True
    Dim attempt As Long, directionOK As Boolean, factorPass As Long, preferPrecise As Boolean
    RPX_HSDObservationBegin auditMode
    hTau = 1#: hKappa = 1#
    For XIteration = 0 To 149
        If RPX_DiagClock() - solveStarted > solveBudget Then
            XStatus = "NUMERICAL_UNKNOWN"
            RPX_DiagEvent "hsd_budget", "classification=NUMERICAL_UNKNOWN;seconds=" & (RPX_DiagClock() - solveStarted) & ";limit=" & solveBudget & ";iteration=" & XIteration
            err.Raise 5, "RPX_Solve", "HSD_SOLVE_BUDGET_EXCEEDED"
        End If
        XResidual = Residuals(hTau): gap = 0#: objective = 0#: hScalarResidual = hKappa
        For i = 0 To XS - 1
            gap = gap + XSlack(i) * XDual(i): hScalarResidual = hScalarResidual + hOffset(i) * XDual(i)
        Next i
        For i = 0 To XN - 1: objective = objective + XQ(i) * xx(i): Next i
        hScalarResidual = hScalarResidual + objective
        For i = 0 To XM - 1: hScalarResidual = hScalarResidual + XERhs(i) * xy(i): Next i
        mu = (gap + hTau * hKappa) / (XB + 1): hMu = mu
        If hTau > 0.0000000001 Then
            If XPrimalResidual / hTau < 0.00000001 And XDualResidual / hTau < 0.0000001 And gap / hTau ^ 2 < 0.00000001 * (1# + Abs(objective / hTau)) Then
                If TryHsdCandidate("OPTIMAL_POINT", hTau, 0#, objective, 0#, objective / hTau * XQScale, gap / hTau ^ 2 * XQScale) Then
                    XResidual = Residuals(): EmitRefineSummary: StoreWarm "hsd": Exit Sub
                End If
            End If
        End If
        If HCertificate() Then Exit Sub
        If hTau > 0.0000000001 Then
            RPX_Stage "HSD " & XIteration & " residual=" & Format$(XResidual / hTau, "0.0E+00") & " gap=" & Format$(gap / hTau ^ 2, "0.0E+00")
        Else
            RPX_Stage "HSD " & XIteration & " tau=" & Format$(hTau, "0.0E+00") & " certificate=" & Format$(hCertificateError, "0.0E+00")
        End If
        hSavedX = xx: hSavedY = xy: hSavedSlack = XSlack: hSavedDual = XDual
        hSavedTau = hTau: hSavedKappa = hKappa: hLiveSaved = True
        RPX_DiagBegin dpScaling: Call Scaling(True): RPX_DiagEnd dpScaling: FactorSystem 0.0000000001, True
        hMatrixEpoch = hMatrixEpoch + 1
        For factorPass = 0 To 2
        attempt = factorPass
        If preferPrecise And hBackend = "AUTO_PIVOT" Then
            If factorPass = 0 Then
                attempt = 1
            ElseIf factorPass = 1 Then
                attempt = 0
            End If
        End If
        hAttempt = attempt: hDirectionPresent = False: hConeDirectionPresent = False
        hMua = 0#: hSigma = 0#: hTarget = 0#: hChi = 0#: hScalarRhs = 0#: hDiagonal = 0#
        hPivot = (hBackend = "PIVOT_ONLY" Or attempt = 2)
        hPreciseFactor = (Not hPivot And attempt = 1)
        hFailure = "": Erase hColumnSolved: Erase dx: Erase dy: Erase ds: Erase dz
        If Not HFactor() Then GoTo HRetry
        If Not HArrow() Then GoTo HRetry
        For i = 0 To XS - 1: rc(i) = XDual(i): Next i
        hRole = "affine"
        If Not HDirection(hKappa) Then GoTo HRetry
        alpha = HStep(): mua = (hTau + alpha * dTau) * (hKappa + alpha * dKappa)
        For i = 0 To XS - 1: mua = mua + (XSlack(i) + alpha * ds(i)) * (XDual(i) + alpha * dz(i)): Next i
        sigma = mua / ((XB + 1) * mu): If sigma > 1# Then sigma = 1#
        If sigma < 0# Then sigma = 0#
        target = mu * sigma ^ 3: correction = hKappa - target / hTau + dTau * dKappa / hTau
        hMua = mua: hSigma = sigma: hTarget = target
        Corrector target
        hRole = "corrector": hDirectionPresent = False
        If Not HDirection(correction) Then GoTo HRetry
        Exit For
HRetry:
        RPX_DiagEvent "hsd_attempt_reject", "iteration=" & XIteration & ";attempt=" & attempt & ";reason=" & hFailure
        If hBackend <> "AUTO_PIVOT" Or attempt = 2 Then err.Raise 5, "RPX_Solve", hFailure
        xx = hSavedX: xy = hSavedY: XSlack = hSavedSlack: XDual = hSavedDual
        hTau = hSavedTau: hKappa = hSavedKappa
        Next factorPass
        preferPrecise = (hPreciseFactor And hBackend = "AUTO_PIVOT")
        alpha = 0.99 * HStep()
        HContext
        If alpha < 0.000000000001 Then err.Raise 5, , "HSD_STEP_STALLED"
        HCommit alpha
        Call RPX_HLUPark
    Next XIteration
    err.Raise 5, , "HSD_NOT_CONVERGED"
End Sub
Public Function RPX_TestHSD(ByVal inputPath As String) As String
    On Error GoTo Failed
    RPX_ReadConic inputPath: Call RPX_OptimizeHSD
    RPX_TestHSD = "PASS " & XStatus & " objective=" & XObjective: Exit Function
Failed:
    RPX_TestHSD = "FAIL " & err.description
End Function
Public Function RPX_TestConic(ByVal inputPath As String, ByVal outputPath As String) As String
    Dim f As Integer, i As Long, started As Double
    On Error GoTo Failed
    started = Timer: XLogPath = outputPath & ".log"
    RPX_ReadConic inputPath: RPX_Optimize
    f = FreeFile: Open outputPath For Output As #f
    Print #f, XObjective; ","; XResidual; ","; XGap; ","; XIteration; ","; Timer - started
    For i = 0 To XN - 1: Print #f, xx(i): Next i
    Close #f: Application.StatusBar = False
    RPX_TestConic = "PASS objective=" & CStr(XObjective): Exit Function
Failed:
    RPX_TestConic = "FAIL " & err.number & " " & err.description: Application.StatusBar = False
End Function
Public Sub RPX_Optimize(Optional ByVal prepared As Boolean = False, Optional ByVal auditMode As String = "", Optional ByVal compiledReady As Boolean = False)
    Dim errorNumber As Long, errorText As String, errorSource As String, errorLine As Long
    On Error GoTo Failed
    XStatus = "RUNNING": RPX_AuditRejected = False
    XIteration = 0: XObjective = 0#: XGap = 0#: XPrimalResidual = 0#: XDualResidual = 0#
    If Not solveEngineAnnounced Then
        solveEngineAnnounced = True
        RPX_DiagEvent "solve_engine", "engine=20260925_warm_off;ver=20260925_2023"
    End If
    RPX_P4Begin
    RPX_SchurReset (auditMode = "lower"), prepared
    ' Compiled arrays are ready, but the original Schur lifetime flag is retained.
    If compiledReady Then prepared = True
    Call OptimizeCore(prepared, auditMode)
    XStatus = "OPTIMAL"
    RPX_P4Summary
    Exit Sub
Failed:
    errorNumber = err.number: errorText = err.description: errorSource = err.source: errorLine = Erl
    RPX_SchurRelease
    RPX_ErrorEvidence errorNumber, errorSource, errorText, errorLine, "optimize"
    XStatus = RPX_DiagFailure(errorNumber, errorText)
    err.Raise errorNumber, errorSource, errorText
End Sub
Public Sub RPX_OptimizeHSD(Optional ByVal prepared As Boolean = False, Optional ByVal auditMode As String = "")
    Dim errorNumber As Long, errorText As String, errorSource As String, errorLine As Long
    On Error GoTo Failed
    XStatus = "RUNNING": RPX_AuditRejected = False
    XIteration = 0: XObjective = 0#: XGap = 0#: XPrimalResidual = 0#: XDualResidual = 0#
    RPX_P4Begin
    RPX_SchurRelease
    RPX_SchurReset
    Call OptimizeHSDCore(prepared, auditMode)
    RPX_HLURelease
    Exit Sub
Failed:
    errorNumber = err.number: errorText = err.description: errorSource = err.source: errorLine = Erl
    RPX_SchurRelease
    RPX_HLURelease
    RPX_ErrorEvidence errorNumber, errorSource, errorText, errorLine, "optimize"
    XStatus = RPX_DiagFailure(errorNumber, errorText)
    err.Raise errorNumber, errorSource, errorText
End Sub

' Right-preconditioned GMRES correction using the existing LDL factor.
' Acceptance uses the existing DD operator and 1e-11 scaled/original gate.
Private Function HKNorm(ByRef values() As Double) As Double
    Dim i As Long, largest As Double, total As Double, value As Double
    For i = LBound(values) To UBound(values)
        If Not RPX_SchurFinite(values(i)) Then err.Raise 5, , "HSD_NONFINITE_KRYLOV"
        If Abs(values(i)) > largest Then largest = Abs(values(i))
    Next i
    If largest = 0# Then Exit Function
    For i = LBound(values) To UBound(values)
        value = values(i) / largest: total = total + value * value
    Next i
    HKNorm = largest * Sqr(total)
End Function
Private Function HKRepair(ByRef bdd() As RPX_DD, ByRef answer() As RPX_DD, ByRef bestError As Double) As Boolean
    Const width As Long = 24
    Const restarts As Long = 3
    Dim basis() As Double, preconditioned() As Double, hess() As Double
    Dim cosine(0 To width - 1) As Double, sine(0 To width - 1) As Double, rhs(0 To width) As Double
    Dim coefficients(0 To width - 1) As Double, residual() As Double, vector() As Double, solved() As Double
    Dim vectorDD() As RPX_DD, product() As RPX_DD, candidate() As RPX_DD, saved() As RPX_DD, term As RPX_DD
    Dim n As Long, i As Long, j As Long, k As Long, pass As Long, cycle As Long, used As Long, steps As Long
    Dim beta As Double, value As Double, dotValue As Double, correction As Double, adjusted As Double, nextValue As Double
    Dim magnitude As Double, currentError As Double, nextNorm As Double, previousError As Double
    n = XN + XM: saved = answer: previousError = bestError
    RPX_CapacityCheck "hsd_krylov", 8# * CDbl(n) * (2# * (width + 1) + 12#)
    ReDim basis(0 To n - 1, 0 To width): ReDim preconditioned(0 To n - 1, 0 To width - 1)
    ReDim hess(0 To width, 0 To width - 1): ReDim residual(0 To n - 1): ReDim vector(0 To n - 1): ReDim vectorDD(0 To n - 1)
    For cycle = 0 To restarts - 1
        Call HContext: HOperator answer, product
        currentError = 0#
        For i = 0 To n - 1
            term = HDSub(bdd(i), product(i)): residual(i) = HDValue(term)
            value = Abs(residual(i)): If value > currentError Then currentError = value
            value = Abs(residual(i) / XKScale(i)): If value > currentError Then currentError = value
        Next i
        If currentError < bestError Then bestError = currentError: saved = answer
        If bestError <= 0.00000000001 Then HKRepair = True: Exit For
        beta = HKNorm(residual): If beta = 0# Then Exit For
        For i = 0 To n - 1: basis(i, 0) = residual(i) / beta: Next i
        For i = 0 To width: rhs(i) = 0#: Next i
        rhs(0) = beta: used = 0
        For j = 0 To width - 1
            HContext
            For i = 0 To n - 1: vector(i) = basis(i, j): Next i
            HBacksolve vector, solved
            For i = 0 To n - 1
                If Not RPX_SchurFinite(solved(i)) Then err.Raise 5, , "HSD_NONFINITE_KRYLOV"
                preconditioned(i, j) = solved(i): vectorDD(i) = HDD(solved(i))
            Next i
            HOperator vectorDD, product
            For i = 0 To n - 1: vector(i) = HDValue(product(i)): Next i
            For k = 0 To j + 1: hess(k, j) = 0#: Next k
            ' Two modified Gram-Schmidt passes; compensated dot products.
            For pass = 0 To 1
                For k = 0 To j
                    dotValue = 0#: correction = 0#
                    For i = 0 To n - 1
                        adjusted = vector(i) * basis(i, k) - correction: nextValue = dotValue + adjusted
                        correction = (nextValue - dotValue) - adjusted: dotValue = nextValue
                    Next i
                    hess(k, j) = hess(k, j) + dotValue
                    For i = 0 To n - 1: vector(i) = vector(i) - dotValue * basis(i, k): Next i
                Next k
            Next pass
            nextNorm = HKNorm(vector): hess(j + 1, j) = nextNorm
            If nextNorm > 0# Then
                For i = 0 To n - 1: basis(i, j + 1) = vector(i) / nextNorm: Next i
            End If
            For k = 0 To j - 1
                value = cosine(k) * hess(k, j) + sine(k) * hess(k + 1, j)
                hess(k + 1, j) = -sine(k) * hess(k, j) + cosine(k) * hess(k + 1, j): hess(k, j) = value
            Next k
            magnitude = Abs(hess(j, j)): If Abs(hess(j + 1, j)) > magnitude Then magnitude = Abs(hess(j + 1, j))
            If magnitude = 0# Then Exit For
            value = magnitude * Sqr((hess(j, j) / magnitude) ^ 2 + (hess(j + 1, j) / magnitude) ^ 2)
            cosine(j) = hess(j, j) / value: sine(j) = hess(j + 1, j) / value
            hess(j, j) = value: hess(j + 1, j) = 0#
            rhs(j + 1) = -sine(j) * rhs(j): rhs(j) = cosine(j) * rhs(j)
            used = j + 1: steps = steps + 1
            If nextNorm = 0# Or Abs(rhs(j + 1)) <= 0.00000000000001 * beta Then Exit For
        Next j
        If used = 0 Then Exit For
        For k = used - 1 To 0 Step -1
            value = rhs(k)
            For j = k + 1 To used - 1: value = value - hess(k, j) * coefficients(j): Next j
            If hess(k, k) = 0# Then Exit For
            coefficients(k) = value / hess(k, k)
        Next k
        If k >= 0 Then Exit For
        candidate = answer
        For i = 0 To n - 1
            For j = 0 To used - 1
                term = HDMul(HDD(coefficients(j)), HDD(preconditioned(i, j)))
                candidate(i) = HDAdd(candidate(i), term)
            Next j
        Next i
        answer = candidate
    Next cycle
    HOperator answer, product: currentError = 0#
    For i = 0 To n - 1
        term = HDSub(bdd(i), product(i)): value = Abs(HDValue(term))
        If value > currentError Then currentError = value
        value = value / XKScale(i): If value > currentError Then currentError = value
    Next i
    If currentError < bestError Then bestError = currentError: saved = answer
    answer = saved: HKRepair = (bestError <= 0.00000000001)
    RPX_DiagEvent "hsd_krylov_repair", "iteration=" & XIteration & ";steps=" & steps & ";before=" & previousError & ";after=" & bestError & ";accepted=" & HKRepair & ";factor_reused=True"
End Function

Private Function HDDFactor() As Boolean
    Dim values() As RPX_DD, work(0 To 2) As RPX_DD, coefficient As RPX_DD, term As RPX_DD, total As RPX_DD
    Dim i As Long, j As Long, k As Long, r As Long, q As Long, p As Long, col As Long, pat As Long, n As Long
    Dim at As Long, workSh(0 To 2) As Double, workSl(0 To 2) As Double
    n = XN + XM: HEnsureDD n
    RPX_CapacityCheck "dd_kkt", 16# * (UBound(XKVal) + 1#)
    ReDim values(0 To UBound(XKVal))
    For p = 0 To UBound(XKVal): values(p) = HDD(XKBase(p)): Next p
    For i = 0 To XB - 1
        If i Mod 256 = 0 Then
            RPX_AssertInputs "hsd_dd_factor"
            If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
            If hSolveEpoch <> P4Solve Or hInputEpoch <> RPX_InputEpoch Then err.Raise 5, , "HSD_CONTEXT_EPOCH_MISMATCH"
            If Not HLiveMatches() Then err.Raise 5, , "HSD_ITERATE_CONTEXT_MISMATCH"
        End If
        pat = XC(i).Pattern
        For j = 0 To XC(i).count - 1
            For r = 0 To XC(i).dimn - 1
                work(r) = HDD(0#)
                For q = 0 To XC(i).dimn - 1
                    at = hCoefAt(i) + q * XC(i).count + j
                    coefficient = HDD(hCoefDD(at).hi)
                    HDMulAddKnown work(r), hHhDD(i, r, q), coefficient, hCoefDD(at).sh, hCoefDD(at).sl
                Next q
                HDSplit work(r).hi, workSh(r), workSl(r)
            Next r
            For k = 0 To j
                total = HDD(0#)
                For r = 0 To XC(i).dimn - 1
                    at = hCoefAt(i) + r * XC(i).count + k
                    HDMulAddKnown total, hCoefDD(at), work(r), workSh(r), workSl(r)
                Next r
                p = XC(pat).slots(j, k): values(p) = HDAdd(values(p), total)
            Next k
        Next j
    Next i
    For i = 0 To n - 1
        p = XKDiag(i)
        If XPerm(i) < XN Then
            values(p) = HDAdd(values(p), HDD(0.0000000001))
        Else
            values(p) = HDD(-0.0000000001)
        End If
    Next i
    For col = 0 To n - 1
        For p = XKPtr(col) To XKPtr(col + 1) - 1
            values(p) = HDMulCoef(hScaleDD(col), values(p))
            values(p) = HDMulCoef(hScaleDD(XKId(p)), values(p))
        Next p
    Next col
    HDDFactor = RPX_DDNumeric(XKPtr, XKId, values)
    RPX_DiagEvent "hsd_dd_ldl", "iteration=" & XIteration & ";nnz=" & RPX_FactorNNZ & ";factor_bytes=" & 16# * (CDbl(RPX_FactorNNZ) + 2# * n + UBound(values) + 2#) & ";accepted=" & HDDFactor
End Function
