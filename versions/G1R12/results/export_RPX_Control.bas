Option Explicit

Public RPX_LastError As String
Public RPX_LastFailureNumber As Long, RPX_LastFailureSource As String
Private priorCancelKey As Long
Public Function RPX_TestResultInvalidation() As String
    Dim result As String
    On Error GoTo Failed
    result = RPX_TestBearingWorkflow()
    If Left$(result, 4) <> "PASS" Then err.Raise 5, , result
    RPX_Sheet("要素定義").Range("S9:T9").Value2 = Array(0#, 0#)
    Call RPX_Run
    If Len(RPX_LastError) = 0 Or RPX_ResultCurrent Or Not IsEmpty(RPX_Sheet("解析結果").Range("B4").Value2) Then err.Raise 5, , "STALE_RESULT_NOT_CLEARED"
    RPX_TestResultInvalidation = "PASS failed analysis clears previous result": Exit Function
Failed:
    RPX_TestResultInvalidation = "FAIL " & err.description
End Function
Public Function RPX_TestRigidCoupling() As String
    Dim ws As Worksheet, lower As Double, upper As Double
    On Error GoTo Failed
    Call RPX_Setup: Set ws = RPX_Sheet("要素定義")
    ws.Range("A7:C12,E7:H8,J7:T9").ClearContents
    ws.Range("A7:C7").Value2 = Array(1, 0#, 0#): ws.Range("A8:C8").Value2 = Array(2, 0#, 1#)
    ws.Range("A9:C9").Value2 = Array(3, 1#, 1#): ws.Range("A10:C10").Value2 = Array(4, 1#, 0#)
    ws.Range("A11:C11").Value2 = Array(5, 0#, 0.5): ws.Range("A12:C12").Value2 = Array(6, 1#, 0.5)
    ws.Range("E7:H7").Value2 = Array(1, 1, 0, "1,2,3,4"): ws.Range("E8:H8").Value2 = Array(2, 1, 1, "5,2,3,6")
    ws.Range("J7:T7").Value2 = Array(1, 1, 4, 1, 1, "固定Y", 0, 0#, 0#, 0#, 0#)
    ws.Range("J8:T8").Value2 = Array(2, 1, 2, 3, 0, "剛体", 2, 0#, 0#, 0#, 0#)
    Set ws = RPX_Sheet("材料データ")
    ws.Range("A5:D5").Value2 = Array(1, 1#, 0#, 0#)
    ws.Range("F5:Q5").Value2 = Array(1, 0.5, 0.75, 1, 0, 1, 0#, 0#, 0#, 0#, 0#, 0#)
    ws.Range("F6:Q6").Value2 = Array(2, 0.5, 1#, 1, 0, 1, 0#, 0#, 0#, 0#, -1#, 0#)
    Set ws = RPX_Sheet("設定")
    ws.Range("B4").Value2 = "極限支持力": ws.Range("B5").Value2 = 16: ws.Range("B7").Value2 = 0: ws.Range("B11").Value2 = 2
    ws.Range("B13").Value2 = "2": ws.Range("B14").Value2 = "2"
    Call RPX_Run: If Len(RPX_LastError) > 0 Then err.Raise 5, , RPX_LastError
    lower = RLowerLoad: upper = RUpperLoad
    If Abs(RPX_BearingPressure(lower) - lower) > 0.000001 Then err.Raise 5, , "RIGID_PRESSURE_FAILED"
    RPX_Sheet("材料データ").Range("F6:Q6").ClearContents
    RPX_Sheet("材料データ").Range("P5").Value2 = -1#
    RPX_Sheet("要素定義").Range("P8").Value2 = 1
    ws.Range("B14").Value2 = "1"
    Call RPX_Run: If Len(RPX_LastError) > 0 Then err.Raise 5, , RPX_LastError
    If Abs(lower - RLowerLoad) > 0.00001 Or Abs(upper - RUpperLoad) > 0.00001 Then err.Raise 5, , "RIGID_COUPLING_EQUIVALENCE_FAILED"
    RPX_TestRigidCoupling = "PASS lower=" & RLowerLoad & " upper=" & RUpperLoad: Exit Function
Failed:
    RPX_TestRigidCoupling = "FAIL " & err.description
End Function
Public Function RPX_TestFrictionFs(Optional ByVal logPath As String = "") As String
    On Error GoTo Failed
    Call RPX_Setup
    RPX_ConfigureLegacyDiagnostics True
    XLogPath = logPath
    RPX_Sheet("材料データ").Range("B5").Value2 = 0#
    RPX_Sheet("材料データ").Range("C5").Value2 = 45#
    RPX_Sheet("設定").Range("B5").Value2 = 64
    RPX_Sheet("設定").Range("B7").Value2 = 0
    RPX_Sheet("設定").Range("B11").Value2 = 2
    Call RPX_Run
    If Len(RPX_LastError) > 0 Then err.Raise 5, , RPX_LastError
    If RLowerFs <= 0# Or RLowerFs > RUpperFs Then err.Raise 5, , "FRICTION_FS_INVALID"
    RPX_TestFrictionFs = "PASS Fs=" & RLowerFs & "," & RUpperFs: Exit Function
Failed:
    RPX_TestFrictionFs = "FAIL " & err.description
End Function
Public Function RPX_TestBearingWorkflow() As String
    Dim ws As Worksheet
    On Error GoTo Failed
    Call RPX_Setup
    RPX_ConfigureLegacyDiagnostics
    Set ws = RPX_Sheet("要素定義")
    ws.Range("A7:C12,E7:H7,J7:T9").ClearContents
    ws.Range("A7:C7").Value2 = Array(1, 0#, 0#)
    ws.Range("A8:C8").Value2 = Array(2, 0#, 1#)
    ws.Range("A9:C9").Value2 = Array(3, 1#, 1#)
    ws.Range("A10:C10").Value2 = Array(4, 1#, 0#)
    ws.Range("E7:H7").Value2 = Array(1, 1, 0, "1,2,3,4")
    ws.Range("J7:T7").Value2 = Array(1, 1, 1, 2, 0, "固定X", 0, 0#, 0#, 0#, 0#)
    ws.Range("J8:T8").Value2 = Array(2, 1, 4, 1, 1, "固定Y", 0, 0#, 0#, 0#, 0#)
    ws.Range("J9:T9").Value2 = Array(3, 1, 2, 3, 0, "自由", 0, 0#, -0.25, 0#, -0.5)
    RPX_Sheet("材料データ").Range("A5:D5").Value2 = Array(1, 1#, 0#, 0#)
    Set ws = RPX_Sheet("設定")
    ws.Range("B4").Value2 = "極限支持力": ws.Range("B5").Value2 = 16
    ws.Range("B7").Value2 = 0: ws.Range("B11").Value2 = 2: ws.Range("B13").Value2 = "3"
    Call RPX_Run
    If Len(RPX_LastError) > 0 Then err.Raise 5, , RPX_LastError
    If Abs(RLowerLoad - 3.5) > 0.00001 Or Abs(RUpperLoad - 3.5) > 0.00001 Then err.Raise 5, , "LOAD_MULTIPLIER_FAILED"
    If Abs(RPX_BearingPressure(RLowerLoad) - 2#) > 0.00001 Then err.Raise 5, , "ULTIMATE_PRESSURE_FAILED"
    RPX_TestBearingWorkflow = "PASS multiplier=" & RLowerLoad & "," & RUpperLoad & " pressure=" & RPX_BearingPressure(RLowerLoad)
    Exit Function
Failed:
    RPX_TestBearingWorkflow = "FAIL " & err.description
End Function
Public Function RPX_TestWorkflow(Optional ByVal adaptive As Boolean = False, Optional ByVal targetElements As Long = 16, Optional ByVal cycles As Long = 1, Optional ByVal subdivisions As Long = 2, Optional ByVal tracePath As String = "") As String
    On Error GoTo Failed
    Call RPX_Setup
    XLogPath = tracePath
    RPX_Sheet("設定").Range("B5").Value2 = targetElements
    RPX_Sheet("設定").Range("B7").Value2 = Abs(CLng(adaptive))
    RPX_Sheet("設定").Range("B8").Value2 = cycles
    RPX_Sheet("設定").Range("B11").Value2 = subdivisions
    Call RPX_CreateMesh
    If Len(RPX_LastError) > 0 Then err.Raise 5, , RPX_LastError
    Call RPX_Run
    If Len(RPX_LastError) > 0 Then err.Raise 5, , RPX_LastError
    RPX_TestWorkflow = "PASS Fs=" & CStr(RLowerFs) & "," & CStr(RUpperFs)
    Exit Function
Failed:
    RPX_TestWorkflow = "FAIL " & err.description & " ADAPT=" & CStr(RPX_Setting("ADAPT", -1)) & " cell=" & CStr(RPX_Sheet("設定").Range("B7").Value2)
End Function
Private Sub StartWork()
    If RPX_Busy Then err.Raise 5, , "解析またはメッシュ生成を実行中です。"
    RPX_Busy = True: XCancel = False: RPX_LastError = "": RPX_ResultCurrent = False
    RPX_LastFailureNumber = 0: RPX_LastFailureSource = ""
    RAdaptRoute = ""
    priorCancelKey = Application.EnableCancelKey: Application.EnableCancelKey = xlErrorHandler
    RPX_InputEnd
    RPX_AffineEnd
    RPX_BearingInvalidate
    RPX_InternalWrite = True
    RPX_EnsureUpperPilotSetting
    Call RPX_C0EnsureSetting
    RPX_C0Ran = False: RPX_C0Adopted = "": RPX_C0TrialLog = ""
    RPX_InternalWrite = False
    RPX_InputBegin
    RPX_FsBracketReset
    RPX_LowerCacheReset
    RPX_SchurRelease
    Call RPX_CertBeginRun
End Sub
Private Sub EndWork()
    RPX_SchurRelease
    RPX_LowerCacheReset
    RPX_InputEnd
    RPX_TestDisarm
    RPX_Busy = False: RPX_InternalWrite = False
    On Error Resume Next
    Application.StatusBar = False
    err.Clear
    On Error GoTo 0
    Application.EnableCancelKey = priorCancelKey
End Sub
Public Sub RPX_CreateMesh()
    Dim failureLine As Long
    If RPX_Busy Then Exit Sub
    On Error GoTo Failed
    Call RPX_Invalidate
    Call StartWork: RPX_DiagStart "mesh": Call RPX_StageConfigure: Call DiagnosticLink
    Call RPX_ReadInputs: Call RPX_GenerateMesh
    Call RPX_WriteMesh: Call RPX_Draw
    RPX_Sheet("解析結果").Range("A1").Value2 = "メッシュ生成済み・未解析"
    RPX_DiagFinish "MESH_READY": Call DiagnosticLink
    Call EndWork: Exit Sub
Failed:
    RPX_LastFailureNumber = err.number: RPX_LastError = err.description: RPX_LastFailureSource = err.source: failureLine = Erl
    RPX_ErrorEvidence RPX_LastFailureNumber, RPX_LastFailureSource, RPX_LastError, failureLine, "control_mesh"
    RPX_DiagFinish RPX_DiagFailure(RPX_LastFailureNumber, RPX_LastError), RPX_LastError: Call DiagnosticLink
    RPX_Sheet("解析結果").Range("A1").Value2 = "メッシュ生成失敗：" & RPX_LastError
    Call EndWork
End Sub
Public Sub RPX_Run()
    Dim mode As String, lower As Double, upper As Double, started As Double, ws As Worksheet, item As Variant, row As Long
    Dim profile As RPX_ProblemProfile
    Dim failureLine As Long, targetMet As Boolean
    If RPX_Busy Then Exit Sub
    On Error GoTo Failed
    Call StartWork: RPX_DiagStart "analysis": RPX_BenchInputCompare: Call RPX_StageConfigure: started = Timer
    RPX_C0SplitMesh = False: RPX_C0LowerElements = 0: RPX_C0UpperElements = 0
    Set ws = RPX_Sheet("解析結果"): ws.cells.ClearContents: ws.Range("A1").Value2 = "解析中"
    Call DiagnosticLink
    RPX_InternalWrite = True
    RPX_EnsureUpperPilotSetting
    RPX_InternalWrite = False
    Call RPX_ReadInputs
    mode = CStr(RPX_Setting("MODE", "Fs"))
    RPX_DiagEvent "settings", "mode=" & mode & ";target=" & DTarget & ";adapt=" & RPX_Setting("ADAPT", 1) & ";cycles=" & RPX_Setting("CYCLES", 3) & ";fs_tolerance=" & RFsTolerance & ";gravity=" & RGravityMode & ";gap_target=" & RPX_Setting("GAP", 0.01) & ";refine=" & RPX_Setting("REFINE", 0.2) & ";upper_pilot=" & RPX_Setting("UPPER_PILOT", 0) & ";upper_pilot_rigid=" & RPX_Setting("UPPER_PILOT_RIGID", 0) & ";lower_fs_seed=" & RPX_Setting("LOWER_FS_SEED", "LEGACY") & ";lower_seed_width=" & RPX_Setting("LOWER_SEED_WIDTH", 0.02) & ";lower_model_cache=" & RPX_Setting("LOWER_MODEL_CACHE", 0) & ";lower_kkt_backend=" & RPX_Setting("LOWER_KKT_BACKEND", "GENERIC") & ";schur_allow_rigid=" & RPX_Setting("SCHUR_ALLOW_RIGID", 0) & ";schur_max_bytes=" & RPX_Setting("SCHUR_MAX_BYTES", 0) & ";test_injection=" & RPX_Setting("TEST_INJECTION", 0)
    If mode <> "Fs" And mode <> "極限支持力" Then err.Raise 5, , "解析種別はFsまたは極限支持力です。"
    RPX_DiagEvent "control_engine", "engine=20260926_p2s_f01_f07;ver=20260926_0945"
    If mode = "極限支持力" Then RPX_BearingValidateSelection
    If mode = "Fs" And RPX_C0AutoRequested() Then
        Call RPX_C0RunSelected(lower, upper)
    ElseIf CDbl(RPX_Setting("ADAPT", 1)) <> 0# Then
        RPX_RunAdapt mode
        If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
        If Not RPX_AdaptPairComplete() Then err.Raise 5, , "ADAPT_PAIR_INCOMPLETE"
        If mode = "Fs" Then
            lower = RLowerFs: upper = RUpperFs
            If lower <= 0# Or upper <= 0# Or lower > upper Then err.Raise 5, , "FS_PAIR_INCOMPLETE"
        Else
            lower = RLowerLoad: upper = RUpperLoad
        End If
        ' G1R12/#1: write geometry and velocity from the same accepted upper witness.
        Call RPX_PrepareUpperOutput(True)
        If mode = "極限支持力" Then RPX_BearingMeasure CStr(RPX_Setting("BEARING_FACES", ""))
        Call RPX_WriteMesh: Call RPX_Draw
    Else
        RAdaptRoute = "nonadaptive"
        Call RPX_GenerateMesh
        RPX_BuildModelProfile profile
        RPX_DiagEvent "problem_profile", RPX_ProfileText(profile)
        RPX_LogFeatureRoutes profile
        Call RPX_WriteMesh: Call RPX_Draw
        If mode = "Fs" Then
            Call RPX_FsPair: lower = RLowerFs: upper = RUpperFs
        Else
            RPX_BearingMeasure CStr(RPX_Setting("BEARING_FACES", ""))
            Call RPX_LoadPair: lower = RLowerLoad: upper = RUpperLoad
        End If
    End If
    RPX_AssertInputs "final_output"
    RPX_TestHook "output"
    RPX_AssertInputs "final_output_after_hook"
    If mode = "極限支持力" Then RPX_BearingAssertCurrent
    targetMet = RPX_TargetMet(mode, lower, upper, True)
    RPX_DiagEvent "lower_layout", RPX_LowerLayoutText()
    RPX_InternalWrite = True
    ws.Range("A1").Value2 = "解析完了"
    ws.Range("A3:C3").Value2 = Array("結果", "下界", "上界")
    If mode = "Fs" Then
        ws.Range("A4").Value2 = "強度低減安全率 Fs [-]"
    Else
        ws.Range("A4").Value2 = "極限荷重倍率 [-]"
        ws.Range("A5:C5").Value2 = Array("平均極限支持圧 [kPa]", RPX_BearingPressure(lower), RPX_BearingPressure(upper))
        ws.Range("A6:C6").Value2 = Array("鉛直極限合力 [kN/m]", RBearingFixed + lower * RBearingReference, RBearingFixed + upper * RBearingReference)
        ws.Range("A7:B7").Value2 = Array("集計面の水平投影幅 [m]", RBearingWidth)
    End If
    ws.Range("B4:C4").Value2 = Array(lower, upper)
    ws.Range("A9:B9").Value2 = Array("区間幅 / 中点", 2# * (upper - lower) / (Abs(upper) + Abs(lower)))
    ws.Range("B9").NumberFormat = "0.00%"
    ws.Range("A10:B10").Value2 = Array("要素数", re)
    If RPX_C0Ran Then
        ws.Range("A2").Value2 = "c=0自動メッシュ: " & RPX_C0Adopted
        ws.Range("A10:C10").Value2 = Array("要素数（下界 / 上界）", RPX_C0LowerElements, RPX_C0UpperElements)
        ws.Range("A19").Value2 = "下界は粗いメッシュ。上界の層はc=0の土だけ。c>0と剛体には入れない。"
        ws.Range("A20").Value2 = RPX_C0TrialLog
        If RPX_CertCurrent("upper") Then
            ws.Range("A14").Value2 = "上界の証明: " & RPX_CertUpper.proofKind
            If RPX_CertUpper.proofKind = "TARGET_WITNESS" Then
                ws.Range("A12:B12").Value2 = Array("上界端点の探索", "未実施。目標区間幅は達成")
                ws.Range("A21").Value2 = "旧上界探索のブラケット幅は、この上界の端点幅ではない: " & CStr(RFsLargestRootBracket)
            End If
        End If
    ElseIf CDbl(RPX_Setting("ADAPT", 1)) <> 0# Then
        ws.Range("A10:C10").Value2 = Array("要素数", RBestLower.elements, RBestUpper.elements)
        ws.Range("A15:C15").Value2 = Array("採用メッシュ", RBestLower.source, RBestUpper.source)
        ws.Range("A2").Value2 = "adaptive mesh: " & RAdaptRoute
        ws.Range("A18:J18").Value2 = Array("Stage", "Elements", "Nodes", "Lower", "Upper", "PairGap", "BestLower", "BestUpper", "FsCalls", "Action")
        row = 19
        For Each item In RAdaptHistory
            ws.cells(row, 1).Resize(1, 10).Value2 = item: row = row + 1
        Next item
    End If
    ws.Range("A11:B11").Value2 = Array("計算時間 [秒]", (Timer - started + 86400#) Mod 86400#)
    If mode = "Fs" Then
        If Not (RPX_CertUpper.valid And RPX_CertUpper.proofKind = "TARGET_WITNESS") Then
            ws.Range("A12:B12").Value2 = Array("Fs探索の最終ブラケット幅", RFsLargestRootBracket)
        End If
        If RFsSearchIncomplete Then
            ws.Range("A1").Value2 = "上下界を取得（Fs探索許容差は未達）"
            ws.Range("A14").Value2 = "証明済みの端点を保持：" & RFsSearchNote
        End If
    End If
    ws.Range("A13").Value2 = "上下界は現在の解析領域・境界条件に対する値です。領域寸法の影響は別途確認してください。"
    ws.Range("A8:B8").Value2 = Array("解析政策", RPX_AnalysisPolicy())
    If RPX_AnalysisPolicy() = "ASSOCIATED" Then ws.Range("D8").Value2 = "関連モデル：低減後のψ=φ。旧政策とは区別して評価。"
    If RPX_UsesDavis() Then ws.Range("D8").Value2 = "使用材料にψ<φがある。表示の上下界はDavis等価関連モデルに対する値。"
    If Not targetMet Then
        ws.Range("A1").Value2 = "上下界を取得（目標区間幅またはFs探索許容差は未達）"
    ElseIf RPX_CertCurrent("upper") And RPX_CertUpper.proofKind = "TARGET_WITNESS" Then
        ws.Range("A1").Value2 = "解析完了（目標区間幅を達成）"
    End If
    ws.Range("B4:C7").NumberFormat = "0.000000"
    RPX_WriteFields ((Not RPX_C0Ran) And (CDbl(RPX_Setting("ADAPT", 1)) <> 0#))
    If XCancel Then err.Raise 18, , "ANALYSIS_CANCELLED"
    RPX_AssertInputs "result_commit"
    RPX_ResultCurrent = True: Call RPX_DrawVelocity
    If Not targetMet Then
        RPX_DiagFinish "PARTIAL_CERTIFIED", RFsSearchNote
    Else
        RPX_DiagFinish "OPTIMAL"
    End If
    Call DiagnosticLink
    ws.Activate: Call EndWork: Exit Sub
Failed:
    RPX_LastFailureNumber = err.number: RPX_LastError = err.description: RPX_LastFailureSource = err.source: failureLine = Erl
    RProgramKind = "": RPX_ResultCurrent = False: Call RPX_BearingInvalidate
    On Error Resume Next
    RPX_ErrorEvidence RPX_LastFailureNumber, RPX_LastFailureSource, RPX_LastError, failureLine, "control_analysis"
    RPX_DiagFinish RPX_DiagFailure(RPX_LastFailureNumber, RPX_LastError), RPX_LastError
    Set ws = RPX_Sheet("解析結果"): ws.Range("B4:D30").ClearContents
    ws.Range("A1").Value2 = "解析未完了：" & RPX_LastError
    Call DiagnosticLink
    On Error GoTo 0
    Call EndWork
End Sub
Private Sub RPX_DrawVelocityCore()
    Dim e As Long, k As Long, stride As Long, maximum As Double, vx As Double, vy As Double, weight As Double
    Dim px As Double, py As Double, xmin As Double, xmax As Double, ymax As Double, ymin As Double, factor As Double
    Dim values() As Double, ws As Worksheet, arrow As shape
    ReDim values(0 To re - 1, 0 To 3)
    For e = 0 To re - 1
        vx = 0#: vy = 0#: px = 0#: py = 0#
        For k = 0 To 5
            weight = -1# / 9#: If k >= 3 Then weight = 4# / 9#
            vx = vx + weight * RUpperField(12 * e + 2 * k): vy = vy + weight * RUpperField(12 * e + 2 * k + 1)
            If k < 3 Then px = px + RXY(RTri(e, k), 0) / 3#: py = py + RXY(RTri(e, k), 1) / 3#
        Next k
        values(e, 0) = px: values(e, 1) = py: values(e, 2) = vx: values(e, 3) = vy
        If Sqr(vx * vx + vy * vy) > maximum Then maximum = Sqr(vx * vx + vy * vy)
    Next e
    If maximum <= 0# Then Exit Sub
    xmin = RXY(0, 0): xmax = xmin: ymin = RXY(0, 1): ymax = ymin
    For e = 1 To rn - 1
        If RXY(e, 0) < xmin Then xmin = RXY(e, 0)
        If RXY(e, 0) > xmax Then xmax = RXY(e, 0)
        If RXY(e, 1) < ymin Then ymin = RXY(e, 1)
        If RXY(e, 1) > ymax Then ymax = RXY(e, 1)
    Next e
    factor = 780# / (xmax - xmin): If factor * (ymax - ymin) > 460# Then factor = 460# / (ymax - ymin)
    Set ws = RPX_Sheet("図"): stride = re \ 200: If stride < 1 Then stride = 1
    For e = 0 To re - 1 Step stride
        If Sqr(values(e, 2) ^ 2 + values(e, 3) ^ 2) > 0.01 * maximum Then
            px = 30# + factor * (values(e, 0) - xmin): py = 55# + factor * (ymax - values(e, 1))
            Set arrow = ws.Shapes.AddLine(px, py, px + 25# * values(e, 2) / maximum, py - 25# * values(e, 3) / maximum)
            arrow.AlternativeText = "RPX_VELOCITY"
            arrow.line.EndArrowheadStyle = 3: arrow.line.ForeColor.RGB = RGB(178, 47, 55): arrow.line.weight = 1.25
        End If
    Next e
    ws.Range("A3").Value2 = "赤矢印：上界の速度機構（表示長さを正規化、変位量ではありません）"
End Sub
Private Sub RPX_WriteMeshCore()
    Dim data() As Variant, e As Long, j As Long, ws As Worksheet
    RPX_InternalWrite = True
    Set ws = RPX_Sheet("節点データ"): ws.cells.ClearContents
    ws.Range("A1:C1").Value2 = Array("節点ID", "X [m]", "Y [m]")
    ReDim data(1 To rn, 1 To 3)
    For e = 0 To rn - 1
        data(e + 1, 1) = e + 1: data(e + 1, 2) = RXY(e, 0): data(e + 1, 3) = RXY(e, 1)
    Next e
    ws.Range("A2").Resize(rn, 3).Value2 = data
    Set ws = RPX_Sheet("要素データ"): ws.cells.ClearContents
    ws.Range("A1:F1").Value2 = Array("要素ID", "節点1", "節点2", "節点3", "材料内部番号", "剛体内部番号")
    ReDim data(1 To re, 1 To 6)
    For e = 0 To re - 1
        data(e + 1, 1) = e + 1
        For j = 0 To 2: data(e + 1, j + 2) = RTri(e, j) + 1: Next j
        data(e + 1, 5) = RMatId(e) + 1: data(e + 1, 6) = RRigidId(e) + 1
    Next e
    ws.Range("A2").Resize(re, 6).Value2 = data
    RPX_InternalWrite = False: RPX_MeshCurrent = True
End Sub
Public Sub RPX_ShowDiagram()
    ' Saved Excel shapes remain available after reopening without a solve.
    If Not RPX_Busy And RPX_MeshCurrent Then Call RPX_Draw
    RPX_Sheet("図").Activate
End Sub
Private Sub RPX_DrawCore()
    Dim ws As Worksheet, shape As shape, edge As Long, a As Long, b As Long, i As Long
    Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, factor As Double
    If rn < 3 Or re < 1 Or Not RPX_MeshCurrent Then Exit Sub
    If Not RPX_AllowDraw Then Exit Sub
    
    Set ws = RPX_Sheet("図")
    For i = ws.Shapes.count To 1 Step -1: ws.Shapes(i).Delete: Next i
    xmin = RXY(0, 0): xmax = xmin: ymin = RXY(0, 1): ymax = ymin
    For i = 1 To rn - 1
        If RXY(i, 0) < xmin Then xmin = RXY(i, 0)
        If RXY(i, 0) > xmax Then xmax = RXY(i, 0)
        If RXY(i, 1) < ymin Then ymin = RXY(i, 1)
        If RXY(i, 1) > ymax Then ymax = RXY(i, 1)
    Next i
    factor = 780# / (xmax - xmin)
    If factor * (ymax - ymin) > 460# Then factor = 460# / (ymax - ymin)
    For edge = 0 To RNEdge - 1
        a = REdges(edge).a: b = REdges(edge).b
        Set shape = ws.Shapes.AddLine(30# + factor * (RXY(a, 0) - xmin), 55# + factor * (ymax - RXY(a, 1)), 30# + factor * (RXY(b, 0) - xmin), 55# + factor * (ymax - RXY(b, 1)))
        shape.line.ForeColor.RGB = RGB(77, 99, 113): shape.line.weight = 0.4
        If REdges(edge).element(1) < 0 Then
            shape.line.weight = 1.5
            If REdges(edge).kind <> "load" Then shape.line.ForeColor.RGB = RGB(33, 109, 181)
            If Abs(REdges(edge).load0(0)) + Abs(REdges(edge).load0(1)) + Abs(REdges(edge).load1(0)) + Abs(REdges(edge).load1(1)) > 0# Then shape.line.ForeColor.RGB = RGB(208, 78, 45)
        End If
    Next edge
    ws.Range("A1").Value2 = "メッシュ：" & re & "要素 / " & rn & "節点"
    ws.Range("A2").Value2 = "青：拘束、橙：面荷重。座標の縦横比は1:1。"
    If RPX_ResultCurrent Then Call RPX_DrawVelocity
End Sub

Public Sub RPX_WriteMesh()
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    RPX_DiagBegin dpOutput
    Call RPX_WriteMeshCore
    RPX_DiagEnd dpOutput: Exit Sub
Failed:
    errorNumber = err.number: errorText = err.description
    RPX_DiagEnd dpOutput
    err.Raise errorNumber, , errorText
End Sub

Public Sub RPX_Draw()
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    If Not RPX_AllowDraw Then Exit Sub
    
    RPX_DiagBegin dpDraw
    Call RPX_DrawCore
    RPX_DiagEnd dpDraw: Exit Sub
Failed:
    errorNumber = err.number: errorText = err.description
    RPX_DiagEnd dpDraw
    err.Raise errorNumber, , errorText
End Sub

Public Sub RPX_DrawVelocity()
    Dim errorNumber As Long, errorText As String
    On Error GoTo Failed
    
    If Not RPX_AllowDraw Then Exit Sub
    RPX_DiagBegin dpDraw
    Call RPX_DrawVelocityCore
    RPX_DiagEnd dpDraw: Exit Sub
Failed:
    errorNumber = err.number: errorText = err.description
    RPX_DiagEnd dpDraw
    err.Raise errorNumber, , errorText
End Sub

Private Sub DiagnosticLink()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = RPX_Sheet("解析結果")
    ws.Range("A16:B16").ClearContents
    ws.Range("A16").Value2 = "診断ログ"
    If Len(RPX_DiagPath) > 0 Then ws.Range("B16").Value2 = RPX_DiagPath
    If Len(RPX_DiagWarning) > 0 Then ws.Range("C16").Value2 = RPX_DiagWarning
End Sub