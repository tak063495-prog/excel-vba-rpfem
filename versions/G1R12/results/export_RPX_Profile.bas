' version=20260926_0945 engine=20260926_p2s_f01_f07
Option Explicit
' B0 physical classification. RM is not a single-soil test: count in-use soil materials.
' Feature gates that are not implemented return False and do not change the solver.
Public Type RPX_ProblemProfile
    mode As String
    adapt As Long
    cycles As Long
    soilElements As Long
    soilMaterials As Long
    phi0Soils As Long
    positivePhiSoils As Long
    c0Soils As Long
    positiveCSoils As Long
    rigidElements As Long
    boundaryRigids As Long
    rigidInternal As Long
    soilRigid As Long
    rigidRigid As Long
    bonds As Long
    declaredRigids As Long
    freeRigidComponents As Long
    hasRigidLoad As Boolean
    hsd As Boolean
    allSoilPhi0 As Boolean
End Type

Public Sub RPX_ResetProblemProfile(ByRef profile As RPX_ProblemProfile)
    profile.mode = ""
    profile.adapt = 0
    profile.cycles = 0
    profile.soilElements = 0
    profile.soilMaterials = 0
    profile.phi0Soils = 0
    profile.positivePhiSoils = 0
    profile.c0Soils = 0
    profile.positiveCSoils = 0
    profile.rigidElements = 0
    profile.boundaryRigids = 0
    profile.rigidInternal = 0
    profile.soilRigid = 0
    profile.rigidRigid = 0
    profile.bonds = 0
    profile.declaredRigids = 0
    profile.freeRigidComponents = 0
    profile.hasRigidLoad = False
    profile.hsd = False
    profile.allSoilPhi0 = False
End Sub

Public Sub RPX_BuildModelProfile(ByRef profile As RPX_ProblemProfile)
    Dim e As Long, f As Long, i As Long, id As Long, rid As Long, j As Long, mat As Long
    Dim key As Variant
    Dim used As Object
    RPX_ResetProblemProfile profile
    profile.mode = CStr(RPX_Setting("MODE", "Fs"))
    profile.adapt = CLng(RPX_Setting("ADAPT", 1))
    profile.cycles = CLng(RPX_Setting("CYCLES", 3))
    profile.declaredRigids = RR
    Set used = CreateObject("Scripting.Dictionary")
    If re > 0 Then
        For e = 0 To re - 1
            If RRigidId(e) < 0 Then
                profile.soilElements = profile.soilElements + 1
                mat = RMatId(e)
                If Not used.Exists(CStr(mat)) Then used(CStr(mat)) = mat
            Else
                profile.rigidElements = profile.rigidElements + 1
            End If
        Next e
    End If
    profile.soilMaterials = used.count
    For Each key In used.keys
        id = CLng(used(key))
        If id >= 0 And id < RM Then
            If RMaterials(id).cohesion <= 0# Then
                profile.c0Soils = profile.c0Soils + 1
            Else
                profile.positiveCSoils = profile.positiveCSoils + 1
            End If
            If RMaterials(id).friction = 0# Then
                profile.phi0Soils = profile.phi0Soils + 1
            ElseIf RMaterials(id).friction > 0# Then
                profile.positivePhiSoils = profile.positivePhiSoils + 1
            End If
        End If
    Next key
    If profile.soilMaterials > 0 And profile.positivePhiSoils = 0 Then profile.allSoilPhi0 = True
    If profile.c0Soils > 0 Then profile.hsd = True
    If RR > 0 Then
        For rid = 0 To RR - 1
            For j = 0 To 2
                If Not RRigids(rid).fixed(j) Then profile.freeRigidComponents = profile.freeRigidComponents + 1
                If RRigids(rid).load0(j) <> 0# Or RRigids(rid).load1(j) <> 0# Then profile.hasRigidLoad = True
            Next j
        Next rid
    End If
    If RNEdge > 0 Then
        For i = 0 To RNEdge - 1
            If REdges(i).element(1) < 0 Then
                If REdges(i).kind = "rigid" Then profile.boundaryRigids = profile.boundaryRigids + 1
            Else
                e = REdges(i).element(0)
                f = REdges(i).element(1)
                If e >= 0 And f >= 0 Then
                    If RRigidId(e) >= 0 And RRigidId(e) = RRigidId(f) Then
                        profile.rigidInternal = profile.rigidInternal + 1
                    Else
                        If RRigidId(e) >= 0 Or RRigidId(f) >= 0 Or RMatId(e) <> RMatId(f) Then profile.bonds = profile.bonds + 1
                        If RRigidId(e) >= 0 And RRigidId(f) >= 0 Then
                            profile.rigidRigid = profile.rigidRigid + 1
                        ElseIf RRigidId(e) >= 0 Or RRigidId(f) >= 0 Then
                            profile.soilRigid = profile.soilRigid + 1
                        End If
                    End If
                End If
            End If
        Next i
    End If
End Sub

Public Function RPX_ProfileText(ByRef profile As RPX_ProblemProfile) As String
    Dim rigidLoad As String, hsdText As String
    If profile.hasRigidLoad Then rigidLoad = "1" Else rigidLoad = "0"
    If profile.hsd Then hsdText = "1" Else hsdText = "0"
    RPX_ProfileText = "mode=" & profile.mode & ";adapt=" & profile.adapt & ";cycles=" & profile.cycles _
        & ";soil_elements=" & profile.soilElements & ";soil_materials=" & profile.soilMaterials _
        & ";phi0_soils=" & profile.phi0Soils & ";positive_phi_soils=" & profile.positivePhiSoils _
        & ";c0_soils=" & profile.c0Soils & ";positive_c_soils=" & profile.positiveCSoils _
        & ";rigid_elements=" & profile.rigidElements & ";boundary_rigids=" & profile.boundaryRigids _
        & ";bonds=" & profile.bonds & ";declared_rigids=" & profile.declaredRigids _
        & ";free_rigid_components=" & profile.freeRigidComponents _
        & ";rigid_load=" & rigidLoad & ";hsd=" & hsdText _
        & ";has_phi0_shortcut=" & Abs(CLng(profile.allSoilPhi0)) & ";backend=" & IIf(profile.hsd, "HSD", "NORMAL") _
        & ";rigid_internal=" & profile.rigidInternal & ";soil_rigid=" & profile.soilRigid & ";rigid_rigid=" & profile.rigidRigid _
        & ";rm=" & RM & ";rr=" & RR & ";rn=" & rn & ";re=" & re
End Function

Private Function UpperPilotEligible(ByRef profile As RPX_ProblemProfile, ByRef reason As String) As Boolean
    reason = ""
    If profile.adapt = 0 Then reason = "adapt_disabled": Exit Function
    If profile.cycles < 1 Then reason = "cycles": Exit Function
    If profile.mode <> "Fs" Then reason = "mode": Exit Function
    If profile.soilElements <= 0 Then reason = "mesh": Exit Function
    If profile.declaredRigids > 0 Or profile.rigidElements > 0 Or profile.boundaryRigids > 0 Then reason = "rigid": Exit Function
    If profile.soilMaterials <> 1 Then reason = "materials": Exit Function
    If profile.positiveCSoils <> 1 Or profile.c0Soils > 0 Then reason = "cohesion": Exit Function
    If profile.positivePhiSoils <> 1 Or profile.phi0Soils > 0 Then reason = "phi0": Exit Function
    If profile.hsd Then reason = "hsd": Exit Function
    UpperPilotEligible = True
End Function

Public Function RPX_CanUseFeature(ByVal feature As String, ByRef profile As RPX_ProblemProfile, ByRef reason As String) As Boolean
    Dim requested As Long
    Dim seed As String, kkt As String
    reason = ""
    Select Case feature
        Case "UPPER_PILOT"
            requested = CLng(RPX_Setting("UPPER_PILOT", 0))
            If requested = 0 Then reason = "disabled": Exit Function
            If requested <> 1 Then reason = "setting": Exit Function
            If Not UpperPilotEligible(profile, reason) Then Exit Function
            RPX_CanUseFeature = True
        Case "UPPER_PILOT_RIGID"
            requested = CLng(RPX_Setting("UPPER_PILOT_RIGID", 0))
            If requested = 0 Then reason = "disabled": Exit Function
            If requested <> 1 Then reason = "setting": Exit Function
            If CLng(RPX_Setting("UPPER_PILOT", 0)) <> 1 Then reason = "pilot_disabled": Exit Function
            If profile.declaredRigids = 0 Then reason = "no_rigid": Exit Function
            If profile.adapt = 0 Or profile.cycles < 1 Or profile.mode <> "Fs" Then reason = "mode_adapt_cycles": Exit Function
            If profile.soilMaterials <> 1 Or profile.positiveCSoils <> 1 Or profile.positivePhiSoils <> 1 Or profile.hsd Then reason = "materials": Exit Function
            RPX_CanUseFeature = True
        Case "LOWER_FS_SEED"
            seed = UCase$(Trim$(CStr(RPX_Setting("LOWER_FS_SEED", "LEGACY"))))
            If seed = "LEGACY" Or Len(seed) = 0 Then reason = "legacy": Exit Function
            If seed <> "UPPER_TRIAL" Then reason = "setting": Exit Function
            If Not UpperPilotEligible(profile, reason) Then Exit Function
            RPX_CanUseFeature = True
        Case "LOWER_MODEL_CACHE"
            requested = CLng(RPX_Setting("LOWER_MODEL_CACHE", 0))
            If requested = 0 Then reason = "disabled": Exit Function
            If requested <> 1 Then reason = "setting": Exit Function
            If Not UpperPilotEligible(profile, reason) Then Exit Function
            RPX_CanUseFeature = True
        Case "LOWER_KKT_BACKEND"
            kkt = UCase$(Trim$(CStr(RPX_Setting("LOWER_KKT_BACKEND", "GENERIC"))))
            If kkt = "GENERIC" Or Len(kkt) = 0 Then reason = "generic": Exit Function
            If kkt <> "COMPARE" And kkt <> "AUTO" Then reason = "setting": Exit Function
            If profile.soilElements = 0 Or profile.c0Soils > 0 Or profile.phi0Soils > 0 Or profile.hsd Then reason = "materials": Exit Function
            If profile.declaredRigids > 0 And CLng(RPX_Setting("SCHUR_ALLOW_RIGID", 0)) <> 1 Then reason = "rigid_disabled": Exit Function
            RPX_CanUseFeature = True
        Case "SCHUR_ALLOW_RIGID"
            If CLng(RPX_Setting("SCHUR_ALLOW_RIGID", 0)) = 0 Then reason = "disabled" Else RPX_CanUseFeature = True
        Case Else
            reason = "unknown_feature"
    End Select
End Function

Private Sub LogRoute(ByVal feature As String, ByVal requested As String, ByVal Eligible As Long, ByVal selected As String, ByVal reason As String)
    RPX_DiagEvent "feature_route", "feature=" & feature & ";requested=" & requested & ";eligible=" & Eligible & ";selected=" & selected & ";bypass_reason=" & reason
End Sub

Public Sub RPX_LogFeatureRoutes(ByRef profile As RPX_ProblemProfile)
    Dim reason As String, seed As String, kkt As String
    Dim requested As Long, Eligible As Long
    reason = ""
    Eligible = 0
    If UpperPilotEligible(profile, reason) Then Eligible = 1
    requested = CLng(RPX_Setting("UPPER_PILOT", 0))
    If requested = 0 Then
        LogRoute "UPPER_PILOT", "0", Eligible, "0", "disabled"
    ElseIf requested <> 1 Then
        LogRoute "UPPER_PILOT", CStr(requested), Eligible, "0", "setting"
    ElseIf Eligible = 0 Then
        LogRoute "UPPER_PILOT", "1", 0, "0", reason
    ElseIf profile.cycles < 1 Then
        LogRoute "UPPER_PILOT", "1", 1, "0", "cycles"
    Else
        LogRoute "UPPER_PILOT", "1", 1, "1", ""
    End If
    requested = CLng(RPX_Setting("UPPER_PILOT_RIGID", 0))
    If requested = 0 Then
        LogRoute "UPPER_PILOT_RIGID", "0", 0, "0", "disabled"
    Else
        If RPX_CanUseFeature("UPPER_PILOT_RIGID", profile, reason) Then
            LogRoute "UPPER_PILOT_RIGID", CStr(requested), 1, "1", "geometry_guard_required"
        Else
            LogRoute "UPPER_PILOT_RIGID", CStr(requested), 0, "0", reason
        End If
    End If
    seed = UCase$(Trim$(CStr(RPX_Setting("LOWER_FS_SEED", "LEGACY"))))
    If RPX_CanUseFeature("LOWER_FS_SEED", profile, reason) Then
        LogRoute "LOWER_FS_SEED", seed, 1, "UPPER_TRIAL", "requires_current_upper_and_no_lower_history"
    Else
        LogRoute "LOWER_FS_SEED", seed, 0, "LEGACY", reason
    End If
    If RPX_CanUseFeature("LOWER_MODEL_CACHE", profile, reason) Then
        LogRoute "LOWER_MODEL_CACHE", "1", 1, "1", "exact_lower_structure;cold_start"
    Else
        LogRoute "LOWER_MODEL_CACHE", CStr(RPX_Setting("LOWER_MODEL_CACHE", 0)), 0, "0", reason
    End If
    kkt = UCase$(Trim$(CStr(RPX_Setting("LOWER_KKT_BACKEND", "GENERIC"))))
    If RPX_CanUseFeature("LOWER_KKT_BACKEND", profile, reason) Then
        LogRoute "LOWER_KKT_BACKEND", kkt, 1, kkt, "lower_algebraic_gate_pending"
    Else
        LogRoute "LOWER_KKT_BACKEND", kkt, 0, "GENERIC", reason
    End If
    If CLng(RPX_Setting("SCHUR_ALLOW_RIGID", 0)) = 0 Then
        LogRoute "SCHUR_ALLOW_RIGID", "0", 0, "0", "disabled"
    Else
        LogRoute "SCHUR_ALLOW_RIGID", "1", 1, "1", "retained_reactions;algebraic_gate_pending"
    End If
    If CLng(RPX_Setting("TEST_INJECTION", 0)) <> 0 Then
        LogRoute "TEST_INJECTION", "1", 1, "armed_only", "explicit_test_entry_required"
    End If
End Sub

Private Function ReactionLayoutText() As String
    Dim rec As Variant, idKey As String
    Dim records As Long, uniqueN As Long
    Dim seen As Object
    On Error GoTo Failed
    If RReactions Is Nothing Then
        ReactionLayoutText = "reaction_records=;reaction_unique="
        Exit Function
    End If
    Set seen = CreateObject("Scripting.Dictionary")
    For Each rec In RReactions
        records = records + 1
        idKey = CStr(rec(0))
        If Not seen.Exists(idKey) Then
            seen(idKey) = 1
            uniqueN = uniqueN + 1
        End If
    Next rec
    ReactionLayoutText = "reaction_records=" & records & ";reaction_unique=" & uniqueN
    Exit Function
Failed:
    ReactionLayoutText = "reaction_records=;reaction_unique="
End Function

Private Function StressLayoutText() As String
    Dim e As Long, stress As Long, n As Long
    On Error GoTo Failed
    n = UBound(RStressStart)
    If re > 0 Then
        For e = 0 To re - 1
            If e > n Then Exit For
            If RStressStart(e) >= 0 Then stress = stress + 18
        Next e
    End If
    StressLayoutText = "stress_dof=" & stress
    Exit Function
Failed:
    StressLayoutText = "stress_dof="
End Function

Public Function RPX_LowerLayoutText() As String
    ' Read-only. Does not rebuild the lower model or invent rigid factors.
    RPX_LowerLayoutText = "xn=" & XN & ";xm=" & XM & ";load_id=" & RLoadVariable & ";" & ReactionLayoutText() & ";" & StressLayoutText()
End Function