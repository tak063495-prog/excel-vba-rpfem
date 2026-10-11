Attribute VB_Name = "RPX_Model"
Option Explicit
Public Type RPX_Material
    cohesion As Double
    friction As Double
    dilation As Double
    gamma As Double
End Type
Public Type RPX_Rigid
    center(0 To 1) As Double
    fixed(0 To 2) As Boolean
    load0(0 To 2) As Double
    load1(0 To 2) As Double
End Type
Public Type RPX_Edge
    a As Long
    b As Long
    element(0 To 1) As Long
    loc(0 To 1, 0 To 2) As Long
    length As Double
    tangent(0 To 1) As Double
    normal(0 To 1) As Double
    kind As String
    rigidId As Long
    load0(0 To 1) As Double
    load1(0 To 1) As Double
End Type
Public rn As Long, re As Long, RM As Long, RR As Long, RNEdge As Long
Public RXY() As Double, RTri() As Long, RMatId() As Long, RRigidId() As Long
Public RMaterials() As RPX_Material, RRigids() As RPX_Rigid, REdges() As RPX_Edge
Public RBody0() As Double, RBody1() As Double, RArea() As Double, RGrad() As Double
Public RStrengthFactor As Double, RStressScale As Double, RSubdivisions As Long
Public RBoundary As Object
Public RAnalysisPolicy As String
Public RRawAuditEnabled As Boolean
Private strengthN As Long, strengthFactor As Double, strengthScale As Double, strengthPolicy As String
Private strengthReady() As Boolean, strengthC() As Double, strengthPhi() As Double
Private strengthSin() As Double, strengthCos() As Double, strengthTan() As Double
Private factorKey(0 To 1) As Double, lastSlot As Long
Public RPX_StrengthComputations As Double
Private sourceC() As Double, sourcePhi() As Double, sourcePsi() As Double
Public Sub RPX_SetAnalysisPolicy(ByVal policy As String)
    policy = UCase$(Trim$(policy))
    Select Case policy
        Case "LEGACY_DAVIS_CLIP", "ASSOCIATED", "DAVIS_EQUIVALENT"
        Case Else: err.Raise 5, "RPX_Model", "INVALID_ANALYSIS_POLICY"
    End Select
    If policy <> RAnalysisPolicy Then
        RAnalysisPolicy = policy: strengthN = 0
        Call RPX_FsBracketReset: RPX_LowerCacheReset: RPX_ResetPreparedState
        RProgramKind = ""
    End If
End Sub
Public Function RPX_AnalysisPolicy() As String
    RPX_AnalysisPolicy = RAnalysisPolicy
    If Len(RPX_AnalysisPolicy) = 0 Then RPX_AnalysisPolicy = "LEGACY_DAVIS_CLIP"
End Function

Public Function RPX_EdgeKey(ByVal a As Long, ByVal b As Long) As String
    If a < b Then RPX_EdgeKey = CStr(a) & ":" & CStr(b) Else RPX_EdgeKey = CStr(b) & ":" & CStr(a)
End Function
Public Sub RPX_Topology()
    Dim e As Long, k As Long, l As Long, a As Long, b As Long, id As Long, side As Long, j As Long
    Dim key As String, edges As Object, rec As Variant, det As Double, dx As Double, dy As Double, cx As Double, cy As Double
    Set edges = RPX_Row(): ReDim REdges(0 To 3 * re - 1): RNEdge = 0
    ReDim RArea(0 To re - 1): ReDim RGrad(0 To re - 1, 0 To 2, 0 To 1)
    For e = 0 To re - 1
        For k = 0 To 2
            If RTri(e, k) < 0 Or RTri(e, k) >= rn Then err.Raise 5, , "INVALID_TRIANGLE_NODE"
        Next k
        a = RTri(e, 0): b = RTri(e, 1): id = RTri(e, 2)
        det = (RXY(b, 0) - RXY(a, 0)) * (RXY(id, 1) - RXY(a, 1)) - (RXY(b, 1) - RXY(a, 1)) * (RXY(id, 0) - RXY(a, 0))
        If det <= 0.000000000001 Then err.Raise 5, , "NONPOSITIVE_TRIANGLE"
        RArea(e) = det / 2#
        For k = 0 To 2
            a = RTri(e, (k + 1) Mod 3): b = RTri(e, (k + 2) Mod 3)
            RGrad(e, k, 0) = (RXY(a, 1) - RXY(b, 1)) / det: RGrad(e, k, 1) = (RXY(b, 0) - RXY(a, 0)) / det
            a = RTri(e, k): b = RTri(e, (k + 1) Mod 3): key = RPX_EdgeKey(a, b)
            If Not edges.Exists(key) Then
                id = RNEdge: RNEdge = RNEdge + 1: edges(key) = id
                REdges(id).element(1) = -1: side = 0
                If a < b Then REdges(id).a = a: REdges(id).b = b Else REdges(id).a = b: REdges(id).b = a
            Else
                id = edges(key): side = 1
                If REdges(id).element(1) >= 0 Then err.Raise 5, , "NONMANIFOLD_EDGE"
            End If
            REdges(id).element(side) = e
            If a < b Then
                REdges(id).loc(side, 0) = k: REdges(id).loc(side, 1) = (k + 1) Mod 3
            Else
                REdges(id).loc(side, 0) = (k + 1) Mod 3: REdges(id).loc(side, 1) = k
            End If
            REdges(id).loc(side, 2) = 3 + k
        Next k
    Next e
    For id = 0 To RNEdge - 1
        With REdges(id)
            dx = RXY(.b, 0) - RXY(.a, 0): dy = RXY(.b, 1) - RXY(.a, 1): .length = Sqr(dx * dx + dy * dy)
            .tangent(0) = dx / .length: .tangent(1) = dy / .length
            .normal(0) = .tangent(1): .normal(1) = -.tangent(0): cx = 0#: cy = 0#
            For k = 0 To 2: cx = cx + RXY(RTri(.element(0), k), 0) / 3#: cy = cy + RXY(RTri(.element(0), k), 1) / 3#: Next k
            cx = cx - (RXY(.a, 0) + RXY(.b, 0)) / 2#: cy = cy - (RXY(.a, 1) + RXY(.b, 1)) / 2#
            If cx * .normal(0) + cy * .normal(1) > 0# Then .normal(0) = -.normal(0): .normal(1) = -.normal(1)
            .kind = "load": .rigidId = -1
            If .element(1) < 0 And Not RBoundary Is Nothing Then
                key = RPX_EdgeKey(.a, .b)
                If RBoundary.Exists(key) Then
                    rec = RBoundary(key): .kind = rec(0): .rigidId = rec(1)
                    For j = 0 To 1: .load0(j) = rec(2 + j): .load1(j) = rec(4 + j): Next j
                End If
            End If
        End With
    Next id
End Sub

Public Sub RPX_P2Grad(ByVal e As Long, ByRef bary() As Double, ByRef g() As Double, Optional ByVal bernstein As Boolean = False)
    Dim k As Long, j As Long, l As Long
    For k = 0 To 2
        For j = 0 To 1
            If bernstein Then g(k, j) = 2# * bary(k) * RGrad(e, k, j) Else g(k, j) = (4# * bary(k) - 1#) * RGrad(e, k, j)
        Next j
        l = (k + 1) Mod 3
        For j = 0 To 1
            g(k + 3, j) = (bary(k) * RGrad(e, l, j) + bary(l) * RGrad(e, k, j))
            If bernstein Then g(k + 3, j) = 2# * g(k + 3, j) Else g(k + 3, j) = 4# * g(k + 3, j)
        Next j
    Next k
End Sub
Public Sub RPX_TraceWeights(ByVal s As Double, ByRef w() As Double)
    w(0) = (1# - s) * (1# - 2# * s): w(1) = s * (2# * s - 1#): w(2) = 4# * s * (1# - s)
End Sub
Public Function RPX_UsesDavis() As Boolean
    Dim i As Long
    If RM <= 0 Or RPX_AnalysisPolicy() = "ASSOCIATED" Then Exit Function
    For i = 0 To RM - 1
        If Abs(RMaterials(i).dilation - RMaterials(i).friction) > 0.0000001 Then
            RPX_UsesDavis = True
            Exit Function
        End If
    Next i
End Function
Public Sub RPX_DavisAssociated(ByVal cohesionIn As Double, ByVal phiRad As Double, ByVal psiRad As Double, ByRef cohesionOut As Double, ByRef phiOut As Double)
    Dim eta As Double, den As Double
    If psiRad > phiRad Then psiRad = phiRad
    If psiRad < 0# Then psiRad = 0#
    den = 1# - Sin(psiRad) * Sin(phiRad)
    If den <= 0.000000000001 Then err.Raise 5, , "DAVIS_DENOMINATOR"
    eta = Cos(psiRad) * Cos(phiRad) / den
    cohesionOut = eta * cohesionIn
    phiOut = Atn(eta * Tan(phiRad))
End Sub
Public Sub RPX_StrengthAt(ByVal e As Long, ByVal factor As Double, ByRef cohesion As Double, ByRef phi As Double)
    Dim deg As Double, phiRad As Double, psiRad As Double, m As Long, policy As String, Slot As Long
    If factor <= 0# Or RStressScale <= 0# Then err.Raise 5, , "INVALID_STRENGTH_SCALE"
    m = RMatId(e): policy = RPX_AnalysisPolicy()
    If m < 0 Or m >= RM Then err.Raise 9, "RPX_Model", "STRENGTH_MATERIAL_ID"
    If strengthN <> RM Or strengthScale <> RStressScale Or strengthPolicy <> policy Then
        ReDim strengthReady(0 To RM - 1, 0 To 1): ReDim strengthC(0 To RM - 1, 0 To 1): ReDim strengthPhi(0 To RM - 1, 0 To 1)
        ReDim strengthSin(0 To RM - 1, 0 To 1): ReDim strengthCos(0 To RM - 1, 0 To 1): ReDim strengthTan(0 To RM - 1, 0 To 1)
        ReDim sourceC(0 To RM - 1, 0 To 1): ReDim sourcePhi(0 To RM - 1, 0 To 1): ReDim sourcePsi(0 To RM - 1, 0 To 1)
        factorKey(0) = 0#: factorKey(1) = 0#: lastSlot = 1
        strengthN = RM: strengthFactor = factor: strengthScale = RStressScale: strengthPolicy = policy
    End If
    If factorKey(0) = factor Then
        Slot = 0
    ElseIf factorKey(1) = factor Then
        Slot = 1
    Else
        Slot = 1 - lastSlot: factorKey(Slot) = factor
        Dim Material As Long
        For Material = 0 To RM - 1: strengthReady(Material, Slot) = False: Next Material
    End If
    lastSlot = Slot
    With RMaterials(m)
        If strengthReady(m, Slot) Then
            If sourceC(m, Slot) <> .cohesion Or sourcePhi(m, Slot) <> .friction Or sourcePsi(m, Slot) <> .dilation Then strengthReady(m, Slot) = False
        End If
        If Not strengthReady(m, Slot) Then
            RPX_StrengthComputations = RPX_StrengthComputations + 1#
            deg = Atn(1#) / 45#
            cohesion = .cohesion / (factor * RStressScale)
            phiRad = Atn(Tan(.friction * deg) / factor)
            If policy = "ASSOCIATED" Then
                phi = phiRad
            Else
                psiRad = .dilation * deg
                Call RPX_DavisAssociated(cohesion, phiRad, psiRad, cohesion, phi)
            End If
            strengthC(m, Slot) = cohesion: strengthPhi(m, Slot) = phi
            sourceC(m, Slot) = .cohesion: sourcePhi(m, Slot) = .friction: sourcePsi(m, Slot) = .dilation
            strengthSin(m, Slot) = Sin(phi): strengthCos(m, Slot) = Cos(phi): strengthTan(m, Slot) = Tan(phi)
            strengthReady(m, Slot) = True
        End If
    End With
    cohesion = strengthC(m, Slot): phi = strengthPhi(m, Slot)
End Sub
Public Sub RPX_StrengthTrigAt(ByVal e As Long, ByVal factor As Double, ByRef cohesion As Double, ByRef phi As Double, ByRef sinPhi As Double, ByRef cosPhi As Double, ByRef tanPhi As Double)
    Dim Material As Long
    RPX_StrengthAt e, factor, cohesion, phi
    Material = RMatId(e)
    sinPhi = strengthSin(Material, lastSlot): cosPhi = strengthCos(Material, lastSlot): tanPhi = strengthTan(Material, lastSlot)
End Sub
Public Sub RPX_StrengthTrig(ByVal e As Long, ByRef cohesion As Double, ByRef phi As Double, ByRef sinPhi As Double, ByRef cosPhi As Double, ByRef tanPhi As Double)
    RPX_StrengthTrigAt e, RStrengthFactor, cohesion, phi, sinPhi, cosPhi, tanPhi
End Sub
Public Sub RPX_Strength(ByVal e As Long, ByRef cohesion As Double, ByRef phi As Double)
    Call RPX_StrengthAt(e, RStrengthFactor, cohesion, phi)
End Sub

Public Sub RPX_ReadModel(ByVal path As String)
    Dim f As Integer, i As Long, j As Long, a As Long, b As Long, nbc As Long, kind As String, rigidId As Long
    Dim p0x As Double, p0y As Double, p1x As Double, p1y As Double, fixedValue As Long
    RProgramKind = "": f = FreeFile: Open path For Input As #f
    Input #f, rn, re, RM, RR, nbc, RStressScale, RStrengthFactor, RSubdivisions
    ReDim RXY(0 To rn - 1, 0 To 1): ReDim RTri(0 To re - 1, 0 To 2)
    ReDim RMatId(0 To re - 1): ReDim RRigidId(0 To re - 1): ReDim RMaterials(0 To RM - 1)
    ReDim RBody0(0 To re - 1, 0 To 1): ReDim RBody1(0 To re - 1, 0 To 1): ReDim RRigids(0 To RR)
    Set RBoundary = RPX_Row()
    For i = 0 To RM - 1
        Input #f, RMaterials(i).cohesion, RMaterials(i).friction, RMaterials(i).gamma
        RMaterials(i).dilation = RMaterials(i).friction
    Next i
    For i = 0 To rn - 1: Input #f, RXY(i, 0), RXY(i, 1): Next i
    For i = 0 To re - 1
        Input #f, RTri(i, 0), RTri(i, 1), RTri(i, 2), RMatId(i), RRigidId(i), RBody0(i, 0), RBody0(i, 1), RBody1(i, 0), RBody1(i, 1)
    Next i
    For i = 0 To RR - 1
        Input #f, RRigids(i).center(0), RRigids(i).center(1)
        For j = 0 To 2: Input #f, fixedValue: RRigids(i).fixed(j) = (fixedValue <> 0): Next j
        For j = 0 To 2: Input #f, RRigids(i).load0(j), RRigids(i).load1(j): Next j
    Next i
    For i = 1 To nbc
        Input #f, a, b, kind, rigidId, p0x, p0y, p1x, p1y
        RBoundary(RPX_EdgeKey(a, b)) = Array(kind, rigidId, p0x, p0y, p1x, p1y)
    Next i
    Close #f: Call RPX_Topology
End Sub
