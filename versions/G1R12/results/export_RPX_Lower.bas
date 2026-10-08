Option Explicit
Public RStressStart() As Long, RLoadVariable As Long
Public RReactions As Collection
Public RLowerEqRoles As Collection
Private forceRows() As Object, rigidLoad0() As Double, rigidLoad1() As Double
Private zeroStress() As Boolean
Private Sub LowerEqual(ByVal row As Object, ByVal rhs As Double, ByVal role As String)
    Dim retained As Boolean
    retained = row.count > 0
    RPX_Equal row, rhs
    If retained Then RLowerEqRoles.Add role
End Sub
Private Sub CoupleRigids(ByVal first As Long, ByVal second As Long, ByVal px As Double, ByVal py As Double)
    Dim j As Long, reaction As Long, firstMoment As Double, secondMoment As Double
    For j = 0 To 1
        reaction = RPX_Variable()
        If j = 0 Then firstMoment = -(py - RRigids(first).center(1)): secondMoment = py - RRigids(second).center(1) Else firstMoment = px - RRigids(first).center(0): secondMoment = -(px - RRigids(second).center(0))
        RPX_Add forceRows(first, j), reaction, 1#: RPX_Add forceRows(second, j), reaction, -1#
        RPX_Add forceRows(first, 2), reaction, firstMoment: RPX_Add forceRows(second, 2), reaction, secondMoment
        RReactions.Add Array(reaction, first, j, 1#, firstMoment): RReactions.Add Array(reaction, second, j, -1#, secondMoment)
    Next j
End Sub
Private Sub FindZeroStress()
    Dim id As Long, e As Long, f As Long, k As Long, a As Long, b As Long, changed As Boolean
    ReDim zeroStress(0 To re - 1, 0 To 5)
    For id = 0 To RNEdge - 1
        With REdges(id)
            e = .element(0)
            If .element(1) < 0 And .kind = "load" And RRigidId(e) < 0 Then
                If RMaterials(RMatId(e)).cohesion = 0# And Abs(.load0(0)) + Abs(.load0(1)) + Abs(.load1(0)) + Abs(.load1(1)) = 0# Then
                    For k = 0 To 2: zeroStress(e, .loc(0, k)) = True: Next k
                End If
            End If
        End With
    Next id
    Do
        changed = False
        For id = 0 To RNEdge - 1
            With REdges(id)
                e = .element(0): f = .element(1)
                If f >= 0 Then
                    If RRigidId(e) < 0 And RRigidId(f) < 0 Then
                        If RMaterials(RMatId(e)).cohesion = 0# And RMaterials(RMatId(f)).cohesion = 0# Then
                            For k = 0 To 2
                                a = .loc(0, k): b = .loc(1, k)
                                If zeroStress(e, a) <> zeroStress(f, b) Then
                                    zeroStress(e, a) = True: zeroStress(f, b) = True: changed = True
                                End If
                            Next k
                        End If
                    End If
                End If
            End With
        Next id
    Loop While changed
End Sub
Public Function RPX_StressId(ByVal e As Long, ByVal k As Long, ByVal component As Long) As Long
    RPX_StressId = RStressStart(e) + 3 * k + component
End Function
Public Function RPX_Traction(ByVal e As Long, ByVal k As Long, ByVal component As Long, ByVal nx As Double, ByVal ny As Double) As Object
    Set RPX_Traction = RPX_Row()
    If component = 0 Then
        RPX_Add RPX_Traction, RPX_StressId(e, k, 0), nx: RPX_Add RPX_Traction, RPX_StressId(e, k, 2), ny
    Else
        RPX_Add RPX_Traction, RPX_StressId(e, k, 2), nx: RPX_Add RPX_Traction, RPX_StressId(e, k, 1), ny
    End If
End Function
Private Sub rigidLoad(ByVal rid As Long, ByVal px As Double, ByVal py As Double, ByVal fx0 As Double, ByVal fy0 As Double, ByVal fx1 As Double, ByVal fy1 As Double)
    rigidLoad0(rid, 0) = rigidLoad0(rid, 0) + fx0 / RStressScale: rigidLoad0(rid, 1) = rigidLoad0(rid, 1) + fy0 / RStressScale
    rigidLoad1(rid, 0) = rigidLoad1(rid, 0) + fx1 / RStressScale: rigidLoad1(rid, 1) = rigidLoad1(rid, 1) + fy1 / RStressScale
    rigidLoad0(rid, 2) = rigidLoad0(rid, 2) + ((px - RRigids(rid).center(0)) * fy0 - (py - RRigids(rid).center(1)) * fx0) / RStressScale
    rigidLoad1(rid, 2) = rigidLoad1(rid, 2) + ((px - RRigids(rid).center(0)) * fy1 - (py - RRigids(rid).center(1)) * fx1) / RStressScale
End Sub
Private Sub RigidTraction(ByVal rid As Long, ByVal edgeId As Long, ByRef traction() As Object)
    Dim k As Long, j As Long, w As Double, mx As Double, my As Double, length As Double
    With REdges(edgeId)
        length = .length
        For k = 0 To 2
            If k = 0 Then
                w = 1# / 12#
            ElseIf k = 1 Then
                w = 1# / 4#
            Else
                w = 1# / 6#
            End If
            mx = length * ((RXY(.a, 0) - RRigids(rid).center(0)) / 3# + (RXY(.b, 0) - RXY(.a, 0)) * w)
            my = length * ((RXY(.a, 1) - RRigids(rid).center(1)) / 3# + (RXY(.b, 1) - RXY(.a, 1)) * w)
            For j = 0 To 1: RPX_Axpy forceRows(rid, j), traction(k, j), length / 3#: Next j
            RPX_Axpy forceRows(rid, 2), traction(k, 1), mx: RPX_Axpy forceRows(rid, 2), traction(k, 0), -my
        Next k
    End With
End Sub
Public Sub RPX_AssembleLower()
    Dim sinPhi As Double, cosPhi As Double, tanPhi As Double
    Dim e As Long, f As Long, k As Long, j As Long, l As Long, rid As Long, id As Long, side As Long, variables As Long
    Dim bary(0 To 2) As Double, g(0 To 5, 0 To 1) As Double, c As Double, phi As Double, nx As Double, ny As Double
    Dim row As Object, a As Object, b As Object, traction(0 To 2, 0 To 1) As Object, kind As String, px As Double, py As Double, reaction As Long
    Set RLowerEqRoles = New Collection
    ReDim RStressStart(0 To re - 1)
    For e = 0 To re - 1
        RStressStart(e) = -1
        If RRigidId(e) < 0 Then RStressStart(e) = variables: variables = variables + 18
    Next e
    RLoadVariable = variables: RPX_Begin variables + 1: RPX_Cost RPX_Unit(RLoadVariable, -1#): Set RReactions = New Collection
    Call FindZeroStress
    ReDim forceRows(0 To RR, 0 To 2): ReDim rigidLoad0(0 To RR, 0 To 2): ReDim rigidLoad1(0 To RR, 0 To 2)
    For rid = 0 To RR - 1
        For j = 0 To 2
            Set forceRows(rid, j) = RPX_Row()
            rigidLoad0(rid, j) = RRigids(rid).load0(j) / RStressScale: rigidLoad1(rid, j) = RRigids(rid).load1(j) / RStressScale
        Next j
    Next rid
    For e = 0 To re - 1
        rid = RRigidId(e)
        If rid >= 0 Then
            px = 0#: py = 0#
            For k = 0 To 2: px = px + RXY(RTri(e, k), 0) / 3#: py = py + RXY(RTri(e, k), 1) / 3#: Next k
            rigidLoad rid, px, py, RArea(e) * RBody0(e, 0), RArea(e) * RBody0(e, 1), RArea(e) * RBody1(e, 0), RArea(e) * RBody1(e, 1)
        Else
            For l = 0 To 2
                For k = 0 To 2: bary(k) = 0#: Next k
                bary(l) = 1#: RPX_P2Grad e, bary, g, True
                For j = 0 To 1
                    Set row = RPX_Unit(RLoadVariable, RBody1(e, j) / RStressScale)
                    For k = 0 To 5
                        RPX_Add row, RPX_StressId(e, k, j), g(k, j): RPX_Add row, RPX_StressId(e, k, 2), g(k, 1 - j)
                    Next k
                    LowerEqual row, -RBody0(e, j) / RStressScale, "SOIL_EQUILIBRIUM:" & e & ":" & l & ":" & j
                Next j
            Next l
            RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi
            For k = 0 To 5
                If zeroStress(e, k) Then
                    For j = 0 To 2: LowerEqual RPX_Unit(RPX_StressId(e, k, j)), 0#, "ZERO_STRESS:" & e & ":" & k & ":" & j: Next j
                Else
                Set row = RPX_Unit(RPX_StressId(e, k, 0), -sinPhi): RPX_Add row, RPX_StressId(e, k, 1), -sinPhi
                Set a = RPX_Unit(RPX_StressId(e, k, 0)): RPX_Add a, RPX_StressId(e, k, 1), -1#
                RPX_Cone3 row, a, RPX_Unit(RPX_StressId(e, k, 2), 2#), 2# * c * cosPhi
                XC(XB - 1).owner = e: XC(XB - 1).parameter1 = k
                End If
            Next k
        End If
    Next e
    For id = 0 To RNEdge - 1
        With REdges(id)
            e = .element(0): f = .element(1): side = 0: nx = .normal(0): ny = .normal(1)
            If RRigidId(e) >= 0 And f >= 0 Then
                If RRigidId(f) < 0 Then e = f: f = .element(0): side = 1: nx = -nx: ny = -ny
            End If
            If RRigidId(e) >= 0 Then
                If f >= 0 Then
                    If RRigidId(e) <> RRigidId(f) Then
                        For k = 0 To 1
                            RPX_VelocityPosition e, .loc(0, k), px, py
                            For j = 0 To 1
                                reaction = RPX_Variable()
                                RPX_Add forceRows(RRigidId(e), j), reaction, 1#: RPX_Add forceRows(RRigidId(f), j), reaction, -1#
                                If j = 0 Then
                                    RPX_Add forceRows(RRigidId(e), 2), reaction, -(py - RRigids(RRigidId(e)).center(1))
                                    RPX_Add forceRows(RRigidId(f), 2), reaction, py - RRigids(RRigidId(f)).center(1)
                                    RReactions.Add Array(reaction, RRigidId(e), j, 1#, -(py - RRigids(RRigidId(e)).center(1)))
                                    RReactions.Add Array(reaction, RRigidId(f), j, -1#, py - RRigids(RRigidId(f)).center(1))
                                Else
                                    RPX_Add forceRows(RRigidId(e), 2), reaction, px - RRigids(RRigidId(e)).center(0)
                                    RPX_Add forceRows(RRigidId(f), 2), reaction, -(px - RRigids(RRigidId(f)).center(0))
                                    RReactions.Add Array(reaction, RRigidId(e), j, 1#, px - RRigids(RRigidId(e)).center(0))
                                    RReactions.Add Array(reaction, RRigidId(f), j, -1#, -(px - RRigids(RRigidId(f)).center(0)))
                                End If
                            Next j
                        Next k
                    End If
                End If
                If f < 0 Then
                    rid = RRigidId(e): px = (RXY(.a, 0) + RXY(.b, 0)) / 2#: py = (RXY(.a, 1) + RXY(.b, 1)) / 2#
                    rigidLoad rid, px, py, .length * .load0(0), .length * .load0(1), .length * .load1(0), .length * .load1(1)
                    For k = 0 To 1
                        RPX_VelocityPosition e, .loc(0, k), px, py
                        If .kind = "rigid" And .rigidId <> rid Then CoupleRigids rid, .rigidId, px, py
                        For j = 0 To 1
                            If .kind = "fixed" Or (.kind = "roller_x" And j = 0) Or (.kind = "roller_y" And j = 1) Then
                                reaction = RPX_Variable(): RPX_Add forceRows(rid, j), reaction, 1#
                                If j = 0 Then
                                    RPX_Add forceRows(rid, 2), reaction, -(py - RRigids(rid).center(1))
                                    RReactions.Add Array(reaction, rid, j, 1#, -(py - RRigids(rid).center(1)))
                                Else
                                    RPX_Add forceRows(rid, 2), reaction, px - RRigids(rid).center(0)
                                    RReactions.Add Array(reaction, rid, j, 1#, px - RRigids(rid).center(0))
                                End If
                            End If
                        Next j
                    Next k
                End If
            Else
                For k = 0 To 2
                    For j = 0 To 1: Set traction(k, j) = RPX_Traction(e, .loc(side, k), j, nx, ny): Next j
                Next k
                rid = -1
                If f >= 0 Then
                    rid = RRigidId(f)
                    If rid < 0 Then
                        For k = 0 To 2
                            For j = 0 To 1
                                Set row = RPX_Row(): RPX_Axpy row, traction(k, j), 1#
                                RPX_Axpy row, RPX_Traction(f, .loc(1 - side, k), j, nx, ny), -1#: LowerEqual row, 0#, "TRACTION_CONTINUITY:" & id & ":" & k & ":" & j
                            Next j
                        Next k
                    End If
                Else
                    kind = .kind
                    If kind = "rigid" Then
                        rid = .rigidId: px = (RXY(.a, 0) + RXY(.b, 0)) / 2#: py = (RXY(.a, 1) + RXY(.b, 1)) / 2#
                        rigidLoad rid, px, py, .length * .load0(0), .length * .load0(1), .length * .load1(0), .length * .load1(1)
                    Else
                        For k = 0 To 2
                            For j = 0 To 1
                                If kind = "load" Or (kind = "roller_x" And j = 1) Or (kind = "roller_y" And j = 0) Then
                                    Set row = RPX_Row(): RPX_Axpy row, traction(k, j), 1#: RPX_Add row, RLoadVariable, -.load1(j) / RStressScale
                                    LowerEqual row, .load0(j) / RStressScale, "BOUNDARY:" & id & ":" & k & ":" & j
                                End If
                            Next j
                        Next k
                    End If
                End If
                If rid >= 0 Then RigidTraction rid, id, traction
            End If
        End With
    Next id
    For rid = 0 To RR - 1
        For j = 0 To 2
            If Not RRigids(rid).fixed(j) Then
                Set row = forceRows(rid, j): RPX_Add row, RLoadVariable, -rigidLoad1(rid, j): LowerEqual row, rigidLoad0(rid, j), "RIGID_BALANCE:" & rid & ":" & j
            End If
        Next j
    Next rid
End Sub
Public Sub RPX_RefreshLower()
    Dim sinPhi As Double, cosPhi As Double, tanPhi As Double
    Dim i As Long, j As Long, e As Long, k As Long, component As Long, c As Double, phi As Double
    For i = 0 To XB - 1
        With XC(i)
            e = .owner: k = .parameter1: RPX_StrengthTrig e, c, phi, sinPhi, cosPhi, tanPhi
            .offset(0) = 2# * c * cosPhi
            For j = 0 To .count - 1
                component = .ids(j) - RPX_StressId(e, k, 0)
                If component = 0 Or component = 1 Then .coef(0, j) = -sinPhi Else .coef(0, j) = 0#
            Next j
        End With
    Next i
End Sub

Public Function RPX_TestModel(ByVal inputPath As String, ByVal mode As String, ByVal outputPath As String) As String
    Dim f As Integer, i As Long, started As Double
    On Error GoTo Failed
    started = Timer: XLogPath = outputPath & ".log"
    RPX_ReadModel inputPath
    If mode = "upper" Then Call RPX_AssembleUpper Else Call RPX_AssembleLower
    RPX_Optimize False, mode
    If mode = "upper" Then Call RPX_AuditUpper Else Call RPX_AuditLower
    f = FreeFile: Open outputPath For Output As #f
    Print #f, XObjective; ","; XResidual; ","; XGap; ","; XIteration; ","; Timer - started
    For i = 0 To XN - 1: Print #f, xx(i): Next i
    Close #f: Application.StatusBar = False
    RPX_TestModel = "PASS objective=" & CStr(XObjective): Exit Function
Failed:
    RPX_TestModel = "FAIL " & err.number & " " & err.description: Application.StatusBar = False
End Function
Public Function RPX_TestOrderModel(ByVal inputPath As String, ByVal mode As String, ByVal outputPath As String) As String
    Dim f As Integer, started As Double
    On Error GoTo Failed
    started = Timer: XLogPath = outputPath & ".log": RPX_ReadModel inputPath
    If mode = "upper" Then Call RPX_AssembleUpper Else Call RPX_AssembleLower
    Call RPX_Compile
    f = FreeFile: Open outputPath For Output As #f
    Print #f, XN; ","; XM; ","; RPX_FactorNNZ; ","; Timer - started
    Close #f: Application.StatusBar = False
    RPX_TestOrderModel = "PASS factor nnz=" & RPX_FactorNNZ: Exit Function
Failed:
    RPX_TestOrderModel = "FAIL " & err.number & " " & err.description: Application.StatusBar = False
End Function