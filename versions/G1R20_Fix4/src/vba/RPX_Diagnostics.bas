Attribute VB_Name = "RPX_Diagnostics"
Option Explicit
Public Type RPX_AuditSnapshot
    maximum(1 To 13) As Double
    location(1 To 13) As String
    seen(1 To 13) As Boolean
    mode As String
End Type
Private Const PhaseCount As Long = 18
' Stage-1 telemetry only: no tolerances, model equations or solver iterates changed.
Public RPX_DiagPath As String, RPX_DiagWarning As String
Public RPX_AuditRejected As Boolean
Private active As Boolean, frequency As Currency, started As Double
Private phaseStart(1 To PhaseCount) As Double, phaseDepth(1 To PhaseCount) As Long
Private seconds(1 To PhaseCount) As Double, calls(1 To PhaseCount) As Long
Private markSeconds(1 To PhaseCount) As Double, markCalls(1 To PhaseCount) As Long
Private solveId As Long, candidate As String, solveMode As String, solveFactor As Double, solveStart As Double, inSolve As Boolean
Private auditMaximum(1 To 13) As Double, auditLocation(1 To 13) As String, auditSeen(1 To 13) As Boolean, auditMode As String
Private sampledPeak As Double, loggingSeconds As Double
Private Const SettingsSheetName As String = "設定"
Private Type RPX_ProcessMemory
    cb As Long
    faults As Long
    peakWorking As LongPtr
    working As LongPtr
    quotaPeakPaged As LongPtr
    quotaPaged As LongPtr
    quotaPeakNonPaged As LongPtr
    quotaNonPaged As LongPtr
    pagefile As LongPtr
    peakPagefile As LongPtr
    privateUsage As LongPtr
End Type
Private Type RPX_SystemMemory
    cb As Long
    load As Long
    totalPhysical As Currency
    availablePhysical As Currency
    totalPagefile As Currency
    availablePagefile As Currency
    totalVirtual As Currency
    availableVirtual As Currency
    availableExtended As Currency
End Type
' 64-bit Excel (VBA7) only. Do not keep a non-PtrSafe Declare in this module.
Private Declare PtrSafe Function RPX_GetProcess Lib "kernel32" Alias "GetCurrentProcess" () As LongPtr
Private Declare PtrSafe Function RPX_GetPID Lib "kernel32" Alias "GetCurrentProcessId" () As Long
Private Declare PtrSafe Function RPX_GetMemory Lib "psapi.dll" Alias "GetProcessMemoryInfo" (ByVal process As LongPtr, ByRef counters As RPX_ProcessMemory, ByVal cb As Long) As Long
Private Declare PtrSafe Function RPX_GetSystemMemory Lib "kernel32" Alias "GlobalMemoryStatusEx" (ByRef status As RPX_SystemMemory) As Long
Private Declare PtrSafe Function RPX_QPC Lib "kernel32" Alias "QueryPerformanceCounter" (ByRef value As Currency) As Long
Private Declare PtrSafe Function RPX_QPF Lib "kernel32" Alias "QueryPerformanceFrequency" (ByRef value As Currency) As Long
Public Enum RPX_Phase
    dpAssembly = 1
    dpConvert = 2
    dpOrdering = 3
    dpSymbolic = 4
    dpKKT = 5
    dpLDL = 6
    dpLinearSolve = 7
    dpScaling = 8
    dpAudit = 9
    dpMesh = 10
    dpDraw = 11
    dpOutput = 12
    dpLDLReach = 13
    dpStepLimit = 14
    dpResiduals = 15
    dpCorrector = 16
    dpGProduct = 17
    dpGTAdd = 18
End Enum

Private stepLimitCalls As Long
Private logBuffer As String
Private logPending As Long
Private Const LogFlushLines As Long = 64
Private stepLimitAlpha1 As Long
Private stepLimitBisect As Long
Private interiorCalls As Long
Private interiorCones As Long
Public RPX_AllowDraw As Boolean
Private savedUpdating As Boolean
Private savedEvents As Boolean
Private uiFrozen As Boolean
Private drawPasses As Long
Private Const HoptBucketCount As Long = 18
Private Const HoptStackMax As Long = 32
Public Const HoptLdl As Long = 1
Public Const HoptLuStore As Long = 2
Public Const HoptLuEliminate As Long = 3
Public Const HoptLuCompress As Long = 4
Public Const HoptTriangular As Long = 5
Public Const HoptOperator As Long = 6
Public Const HoptProduct As Long = 7
Public Const HoptBorder As Long = 8
Public Const HoptBackward As Long = 9
Public Const HoptInputRead As Long = 10
Public Const HoptInputCompare As Long = 11
Public Const HoptLiveKey As Long = 12
Public Const HoptModelKey As Long = 13
Public Const HoptStageBar As Long = 14
Public Const HoptStageFile As Long = 15
Public Const HoptStageLog As Long = 16
Public Const HoptStageYield As Long = 17
Public Const HoptInputAssert As Long = 18
Private hoptSp As Long
Private hoptOverflow As Long
Private hoptStack(1 To HoptStackMax) As Long
Private hoptStart(1 To HoptStackMax) As Double
Private hoptChild(1 To HoptStackMax) As Double
Private hoptInclusive(1 To HoptBucketCount) As Double
Private hoptExclusive(1 To HoptBucketCount) As Double
Private hoptCalls(1 To HoptBucketCount) As Long
Private hoptMarkIncl(1 To HoptBucketCount) As Double
Private hoptMarkExcl(1 To HoptBucketCount) As Double
Private hoptMarkCalls(1 To HoptBucketCount) As Long
Private hoptFinite As Double, hoptUpdates As Double, hoptMaxArena As Double, hoptMaxBytes As Double
Private hoptSolveFinite As Double, hoptSolveUpdates As Double, hoptSolveMaxArena As Double, hoptSolveMaxBytes As Double
Private hoptMarkFinite As Double, hoptMarkUpdates As Double
Private hoptLuFactors As Long, hoptMarkFactors As Long, hoptMaxLnnz As Long, hoptMaxUnnz As Long
Private hoptSolveMaxLnnz As Long, hoptSolveMaxUnnz As Long
Private hoptRowScanNodes As Double, hoptSolveRowScanNodes As Double
Private hoptRowScanRows As Double, hoptSolveRowScanRows As Double
Private hoptIndexDirect As Long, hoptSolveIndexDirect As Long
Private hoptIteration As Long, hoptAttempt As Long, hoptMatrix As Long, hoptFactor As Long, hoptRhs As Long

Private Sub FreezeUi()
    If uiFrozen Then Exit Sub
    savedUpdating = Application.ScreenUpdating
    savedEvents = Application.EnableEvents
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    uiFrozen = True
End Sub

Private Sub ThawUi()
    If Not uiFrozen Then Exit Sub
    On Error Resume Next
    Application.ScreenUpdating = savedUpdating
    Application.EnableEvents = savedEvents
    On Error GoTo 0
    uiFrozen = False
End Sub


Public Function RPX_DiagnosticMode() As Long
    Dim raw As Variant, legacy As Variant, ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SettingsSheetName)
    legacy = 1
    If Len(CStr(RPX_InputCell(ws, 15, 4))) = 0 Then
        raw = RPX_InputCell(ws, 15, 2)
        If Not IsError(raw) Then
            If IsNumeric(raw) Then
                If CDbl(raw) = 0# Or CDbl(raw) = 1# Then legacy = raw
            End If
        End If
    End If
    raw = RPX_Setting("DIAGNOSTIC_MODE", legacy)
    If IsError(raw) Then err.Raise 5, "RPX_Diagnostics", "INVALID_SETTING DIAGNOSTIC_MODE"
    If Not IsNumeric(raw) Then err.Raise 5, "RPX_Diagnostics", "INVALID_SETTING DIAGNOSTIC_MODE"
    If CDbl(raw) <> 0# And CDbl(raw) <> 1# Then err.Raise 5, "RPX_Diagnostics", "INVALID_SETTING DIAGNOSTIC_MODE"
    RPX_DiagnosticMode = CLng(raw)
End Function

Public Function RPX_DiagClock() As Double
    RPX_DiagClock = ClockSeconds()
End Function
Private Function ClockSeconds() As Double
    Dim count As Currency
    If frequency <> 0 Then
        If RPX_QPC(count) <> 0 Then ClockSeconds = count / frequency: Exit Function
    End If
    ClockSeconds = CDbl(Date) * 86400# + Timer
End Function
Private Function Clean(ByVal value As String) As String
    Clean = Replace(Replace(Replace(value, vbTab, " "), vbCr, " "), vbLf, " ")
End Function
Private Function PhaseName(ByVal phase As Long) As String
    PhaseName = Choose(phase, "assembly", "convert_pattern", "ordering", "symbolic", "kkt_update", "numeric_ldl", "linear_solve_and_refinement", "nt_scaling", "physical_audit", "mesh", "draw", "sheet_output", "ldl_reach_cache", "step_limit", "residuals", "corrector", "gproduct", "gtadd")
End Function
Private Function LogsPhaseEvent(ByVal phase As RPX_Phase) As Boolean
    Select Case phase
        Case dpAssembly, dpConvert, dpOrdering, dpSymbolic, dpKKT, dpLDL, dpMesh, dpDraw, dpOutput, dpLDLReach
            LogsPhaseEvent = True
    End Select
End Function
Public Sub RPX_DiagInteriorSample(ByVal examined As Long)
    If Not active Then Exit Sub
    interiorCalls = interiorCalls + 1
    interiorCones = interiorCones + examined
End Sub
Public Sub RPX_DiagStepLimitSample(ByVal alphaOne As Boolean, ByVal bisectCount As Long)
    If Not active Then Exit Sub
    stepLimitCalls = stepLimitCalls + 1
    If alphaOne Then
        stepLimitAlpha1 = stepLimitAlpha1 + 1
    Else
        stepLimitBisect = stepLimitBisect + bisectCount
    End If
End Sub
Private Function AuditName(ByVal category As Long) As String
    AuditName = Choose(category, "yield", "element_equilibrium", "internal_traction", "boundary_traction", "rigid_force", "rigid_moment", "flow", "bond_velocity", "jump", "boundary_velocity", "rigid_velocity", "reference_work", "objective_difference")
End Function
Private Function ProcessBytes(ByRef privateBytes As Double, ByRef peakCommit As Double, ByRef workingBytes As Double) As Boolean
    Dim memory As RPX_ProcessMemory
    On Error GoTo Missing
    memory.cb = LenB(memory)
    If RPX_GetMemory(RPX_GetProcess(), memory, memory.cb) = 0 Then Exit Function
    privateBytes = CDbl(memory.privateUsage): peakCommit = CDbl(memory.peakPagefile): workingBytes = CDbl(memory.working)
    If privateBytes < 0 Or peakCommit < 0 Or workingBytes < 0 Then Exit Function
    If privateBytes > sampledPeak Then sampledPeak = privateBytes
    ProcessBytes = True
Missing:
End Function
Private Sub FlushLog()
    Dim fs As Object, stream As Object, logStarted As Double
    If logPending <= 0 Or Len(RPX_DiagPath) = 0 Or Len(logBuffer) = 0 Then Exit Sub
    On Error GoTo Failed
    logStarted = ClockSeconds()
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set stream = fs.OpenTextFile(RPX_DiagPath, 8, False, -1)
    stream.Write logBuffer
    stream.Close
    logBuffer = "": logPending = 0
    loggingSeconds = loggingSeconds + ClockSeconds() - logStarted
    Exit Sub
Failed:
    RPX_DiagWarning = "診断ログを書き込めません：" & err.description
    logBuffer = "": logPending = 0
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    loggingSeconds = loggingSeconds + ClockSeconds() - logStarted
End Sub
Public Sub RPX_DiagFlush()
    FlushLog
End Sub
Public Sub RPX_DiagEvent(ByVal eventName As String, Optional ByVal detail As String = "")
    Dim mem As String, privateBytes As Double, peak As Double, working As Double, line As String
    If Not active Or Len(RPX_DiagPath) = 0 Then Exit Sub
    On Error GoTo Failed
    mem = "NA" & vbTab & "NA" & vbTab & "NA"
    If ProcessBytes(privateBytes, peak, working) Then mem = CStr(privateBytes) & vbTab & CStr(peak) & vbTab & CStr(working)
    line = Format$(now, "yyyy-mm-dd hh:nn:ss") & vbTab & Format$(ClockSeconds() - started, "0.000000") & vbTab & solveId & vbTab & Clean(candidate) & vbTab & solveMode & vbTab & CStr(solveFactor) & vbTab & eventName & vbTab & mem & vbTab & Clean(detail)
    logBuffer = logBuffer & line & vbCrLf
    logPending = logPending + 1
    If logPending >= LogFlushLines Then FlushLog
    Exit Sub
Failed:
    RPX_DiagWarning = "診断ログを書き込めません：" & err.description
End Sub
Public Sub RPX_DiagStart(ByVal operation As String)
    Dim fs As Object, stream As Object, folder As String, base As String, suffix As Long, bits As String, selectedMode As Long
    selectedMode = RPX_DiagnosticMode()
    active = False: RPX_DiagPath = "": RPX_DiagWarning = "": sampledPeak = 0#: loggingSeconds = 0#: logBuffer = "": logPending = 0
    solveId = 0: candidate = "": solveMode = "": solveFactor = 0#: inSolve = False
    Erase seconds: Erase calls: Erase phaseDepth: Erase markSeconds: Erase markCalls
    Erase hoptInclusive: Erase hoptExclusive: Erase hoptCalls: Erase hoptStack: Erase hoptStart: Erase hoptChild
    hoptSp = 0: hoptOverflow = 0
    hoptFinite = 0#: hoptUpdates = 0#: hoptMaxArena = 0#: hoptMaxBytes = 0#
    hoptSolveFinite = 0#: hoptSolveUpdates = 0#: hoptSolveMaxArena = 0#: hoptSolveMaxBytes = 0#
    hoptMarkFinite = 0#: hoptMarkUpdates = 0#: hoptLuFactors = 0: hoptMarkFactors = 0
    hoptMaxLnnz = 0: hoptMaxUnnz = 0: hoptSolveMaxLnnz = 0: hoptSolveMaxUnnz = 0
    hoptRowScanNodes = 0#: hoptSolveRowScanNodes = 0#: hoptRowScanRows = 0#: hoptSolveRowScanRows = 0#
    hoptIndexDirect = 0: hoptSolveIndexDirect = 0
    hoptIteration = 0: hoptAttempt = 0: hoptMatrix = 0: hoptFactor = 0: hoptRhs = 0
    frequency = 0: On Error Resume Next: RPX_QPF frequency: On Error GoTo Failed
    started = ClockSeconds(): active = True
    RPX_AllowDraw = True
    drawPasses = 0
    Call FreezeUi
    If selectedMode = 0 Then
        RPX_DiagPath = ""
        Exit Sub
    End If
    Set fs = CreateObject("Scripting.FileSystemObject")
    If Len(ThisWorkbook.path) = 0 Then err.Raise 5, , "ブックを保存してから計測してください。"
    folder = ThisWorkbook.path & Application.PathSeparator & "RPFEM_logs"
    If Not fs.FolderExists(folder) Then fs.CreateFolder folder
    base = folder & Application.PathSeparator & Format$(now, "yyyymmdd_hhnnss") & "_" & CStr(RPX_GetPID()) & "_" & operation
    RPX_DiagPath = base & ".tsv"
    Do While fs.FileExists(RPX_DiagPath)
        suffix = suffix + 1: RPX_DiagPath = base & "_" & suffix & ".tsv"
    Loop
    Set stream = fs.CreateTextFile(RPX_DiagPath, False, True)
    stream.WriteLine "timestamp" & vbTab & "elapsed_s" & vbTab & "solve_id" & vbTab & "candidate" & vbTab & "mode" & vbTab & "strength_factor" & vbTab & "event" & vbTab & "private_bytes" & vbTab & "process_peak_commit_bytes" & vbTab & "working_set_bytes" & vbTab & "detail"
    stream.Close
#If Win64 Then
    bits = "64"
#Else
    bits = "32"
#End If
    RPX_DiagEvent "run_start", "version=stage1_20260920_stepA_measure;excel_bits=" & bits & ";excel=" & Application.Version & ";operation=" & operation & ";memory_peaks=whole_process_lifetime"
    RPX_DiagEvent "diag_engine", "engine=20260926_log_buffer;ver=20260926_0430;flush_lines=" & LogFlushLines
    FlushLog
    Exit Sub
Failed:
    RPX_DiagWarning = "診断ログを作成できません：" & err.description
    RPX_DiagPath = ""
    On Error Resume Next
    If Not stream Is Nothing Then stream.Close
    Call ThawUi
End Sub
Public Sub RPX_DiagBegin(ByVal phase As RPX_Phase)
    If Not active Then Exit Sub
    If phaseDepth(phase) = 0 Then
        If LogsPhaseEvent(phase) Then RPX_DiagEvent "phase_begin", "phase=" & PhaseName(phase)
        phaseStart(phase) = ClockSeconds()
    End If
    phaseDepth(phase) = phaseDepth(phase) + 1
End Sub
Public Sub RPX_DiagEnd(ByVal phase As RPX_Phase)
    Dim elapsed As Double
    If Not active Or phaseDepth(phase) = 0 Then Exit Sub
    phaseDepth(phase) = phaseDepth(phase) - 1
    If phaseDepth(phase) = 0 Then
        elapsed = ClockSeconds() - phaseStart(phase)
        seconds(phase) = seconds(phase) + elapsed
        calls(phase) = calls(phase) + 1
        If LogsPhaseEvent(phase) Then RPX_DiagEvent "phase_end", "phase=" & PhaseName(phase) & ";seconds=" & elapsed
        If phase = dpDraw Then
            drawPasses = drawPasses + 1
            RPX_AllowDraw = False
        End If
    End If
End Sub
Public Sub RPX_DiagClosePhases()
    Dim i As Long
    For i = 1 To PhaseCount
        If phaseDepth(i) > 0 Then
            RPX_DiagEvent "phase_interrupted", "phase=" & PhaseName(i)
            phaseDepth(i) = 1: RPX_DiagEnd i
        End If
    Next i
End Sub
Public Sub RPX_DiagCandidate(ByVal label As String)
    candidate = label: RPX_DiagEvent "candidate_begin", "elements=" & re & ";nodes=" & rn
End Sub
Public Sub RPX_DiagSolveStart(ByVal mode As String, ByVal factor As Double, ByVal prepared As Boolean)
    Dim i As Long
    If Not active Then Exit Sub
    solveId = solveId + 1: solveMode = mode: solveFactor = factor: solveStart = ClockSeconds(): inSolve = True
    For i = 1 To PhaseCount: markSeconds(i) = seconds(i): markCalls(i) = calls(i): Next i
    stepLimitCalls = 0: stepLimitAlpha1 = 0: stepLimitBisect = 0
    interiorCalls = 0: interiorCones = 0
    Call HoptSnap
    RPX_DiagEvent "solve_start", "prepared=" & prepared & ";elements=" & re & ";nodes=" & rn & ";subdivisions=" & RSubdivisions
End Sub
Public Sub RPX_DiagCounts()
    Dim i As Long, coefficients As Double, scalar As Long, lorentz As Long, eqnnz As Double, knnz As Double
    If Not active Then Exit Sub
    For i = 0 To XB - 1
        coefficients = coefficients + CDbl(XC(i).count) * XC(i).dimn
        If XC(i).dimn = 1 Then scalar = scalar + 1 Else lorentz = lorentz + 1
    Next i
    eqnnz = XEPtr(XM): knnz = XKPtr(XN + XM)
    RPX_DiagEvent "model_counts", "variables=" & XN & ";equalities=" & XM & ";cones=" & XB & ";scalar_cones=" & scalar & ";soc3=" & lorentz & ";cone_coefficients=" & coefficients & ";eq_nnz=" & eqnnz & ";kkt_nnz=" & knnz & ";factor_nnz=" & RPX_FactorNNZ
End Sub
Public Sub RPX_DiagSolveEnd(ByVal state As String, Optional ByVal message As String = "")
    Dim i As Long
    If Not active Or Not inSolve Then Exit Sub
    Call RPX_DiagClosePhases
    Call HoptEmit("solve")
    For i = 1 To PhaseCount
        If calls(i) <> markCalls(i) Then RPX_DiagEvent "solve_phase", "phase=" & PhaseName(i) & ";seconds=" & CStr(seconds(i) - markSeconds(i)) & ";calls=" & CStr(calls(i) - markCalls(i))
    Next i
    RPX_DiagEvent "inner_stats", "step_limit_calls=" & stepLimitCalls & ";step_limit_alpha1=" & stepLimitAlpha1 & ";step_limit_bisections=" & stepLimitBisect & ";interior_calls=" & interiorCalls & ";interior_cones=" & interiorCones
    RPX_DiagEvent "solve_end", "status=" & state & ";wall_seconds=" & CStr(ClockSeconds() - solveStart) & ";last_iteration_index=" & XIteration & ";objective=" & XObjective & ";primal=" & XPrimalResidual & ";dual=" & XDualResidual & ";gap=" & XGap & ";message=" & message
    inSolve = False
    FlushLog
End Sub
Public Function RPX_DiagFailure(ByVal number As Long, ByVal message As String) As String
    RPX_DiagFailure = "NUMERICAL_FAILURE"
    If InStr(message, "AUDIT") > 0 Or (RPX_AuditRejected And InStr(message, "NOT_CONVERGED") > 0) Then RPX_DiagFailure = "AUDIT_FAILED"
    If number = 7 Or InStr(message, "MEMORY") > 0 Or InStr(message, "CAPACITY") > 0 Then RPX_DiagFailure = "CAPACITY_EXCEEDED"
    If number = 18 Or XCancel Then RPX_DiagFailure = "CANCELLED"
End Function
Public Sub RPX_DiagFinish(ByVal state As String, Optional ByVal message As String = "")
    Dim i As Long
    If Not active Then Exit Sub
    If inSolve Then RPX_DiagSolveEnd state, message
    Call RPX_DiagClosePhases
    Call HoptEmit("run")
    RPX_DiagEvent "hopt0_coverage", "numeric_ldl_zero_calls=not_executed;linear_solve_zero_calls=not_executed;previous_zero_meant=unmeasured;hsd_ldl_is=dpLDL;hsd_triangular_is=dpLinearSolve;lu_parts=hopt0;overflow=" & hoptOverflow
    For i = 1 To PhaseCount
        RPX_DiagEvent "run_phase", "phase=" & PhaseName(i) & ";seconds=" & seconds(i) & ";calls=" & calls(i)
    Next i
    RPX_DiagEvent "run_end", "status=" & state & ";solve_count=" & solveId & ";sampled_peak_private_bytes=" & sampledPeak & ";logging_seconds_before_run_end=" & loggingSeconds & ";message=" & message
    FlushLog
    active = False
    RPX_AllowDraw = True
    Call ThawUi
End Sub
Public Sub RPX_DiagAuditStart(ByVal mode As String)
    auditMode = mode: Erase auditMaximum: Erase auditLocation: Erase auditSeen
    RPX_DiagBegin dpAudit
End Sub
Public Sub RPX_DiagAuditCheck(ByVal category As Long, ByVal value As Double, ByVal indexA As Long, Optional ByVal indexB As Long = -1, Optional ByVal indexC As Long = -1)
    If Not active Then Exit Sub
    auditSeen(category) = True
    If value > auditMaximum(category) Then
        auditMaximum(category) = value
        auditLocation(category) = "index_a=" & indexA & ";index_b=" & indexB & ";index_c=" & indexC
    End If
End Sub
Public Sub RPX_DiagAuditEnd(ByVal state As String)
    Dim i As Long, threshold As Double
    RPX_DiagEnd dpAudit
    If Not active Then Exit Sub
    For i = 1 To 13
        threshold = 0.0000001: If i = 13 Then threshold = 0.000002
        If auditSeen(i) Then RPX_DiagEvent "audit", "mode=" & auditMode & ";status=" & state & ";category=" & AuditName(i) & ";max_raw=" & auditMaximum(i) & ";threshold=" & threshold & ";category_pass=" & CStr(auditMaximum(i) <= threshold) & ";location=" & auditLocation(i) & ";field_strength=" & RStrengthFactor & ";in_solve=" & inSolve & ";normalization=legacy_unchanged"
    Next i
End Sub
Public Sub RPX_CapacityCheck(ByVal stage As String, ByVal additionalBytes As Double)
    Dim sys As RPX_SystemMemory, privateBytes As Double, peak As Double, working As Double
    Dim available As Double, limit As Double, systemOK As Boolean, processOK As Boolean
    sys.cb = LenB(sys)
    On Error Resume Next
    systemOK = (RPX_GetSystemMemory(sys) <> 0)
    processOK = ProcessBytes(privateBytes, peak, working)
    On Error GoTo 0
    If additionalBytes < 0# Then err.Raise 7, , "CAPACITY_ESTIMATE_OVERFLOW"
    If Not systemOK Then
        RPX_DiagEvent "capacity_missing", "stage=" & stage & ";estimated_additional_bytes=" & additionalBytes & ";reason=system_memory_unavailable"
        Exit Sub
    End If
    available = CDbl(sys.availablePhysical) * 10000# * 0.65
    If CDbl(sys.availablePagefile) * 10000# * 0.65 < available Then available = CDbl(sys.availablePagefile) * 10000# * 0.65
    limit = CDbl(sys.totalPhysical) * 10000# * 0.65
#If Win64 Then
    If limit < 4294967296# Then limit = 4294967296#
#Else
    If limit > 1288490188# Then limit = 1288490188#
    If CDbl(sys.availableVirtual) * 10000# * 0.65 < available Then available = CDbl(sys.availableVirtual) * 10000# * 0.65
#End If
    If processOK And privateBytes > limit Then limit = privateBytes * 1.15
    RPX_DiagEvent "capacity", "stage=" & stage & ";estimated_additional_bytes=" & additionalBytes & ";available_budget_bytes=" & available & ";process_limit_bytes=" & limit & ";process_memory_available=" & processOK
    If additionalBytes > available Then
        err.Raise 7, , "CAPACITY_MEMORY_BUDGET stage=" & stage & " estimated_additional_MiB=" & Format$(additionalBytes / 1048576#, "0") & " available_budget_MiB=" & Format$(available / 1048576#, "0") & "：密度・細分数または使用可能メモリを確認してください。設定は自動変更していません。"
    End If
    If processOK And privateBytes + additionalBytes > limit And additionalBytes > available * 0.25 Then
        err.Raise 7, , "CAPACITY_MEMORY_BUDGET stage=" & stage & " estimated_additional_MiB=" & Format$(additionalBytes / 1048576#, "0") & " available_budget_MiB=" & Format$(available / 1048576#, "0") & "：密度・細分数または使用可能メモリを確認してください。設定は自動変更していません。"
    End If
End Sub
Public Sub RPX_CapacityAssembly(ByVal mode As String)
    Dim e As Long, zeroPhi As Long, cones As Double, variables As Double, edges As Double, estimate As Double, n As Double
    n = RSubdivisions
    For e = 0 To re - 1
        If RRigidId(e) < 0 Then
            If RMaterials(RMatId(e)).friction = 0# Then zeroPhi = zeroPhi + 1
        End If
    Next e
    edges = RNEdge
    If mode = "upper" Then
        variables = 12# * re + 3# * RR + zeroPhi * (n + 1#) * (n + 2#) / 2#
        cones = 3# * re + zeroPhi * (n + 1#) * (n + 2#) / 2# + edges * 6# * n
        estimate = cones * 2200# + variables * 2400# + edges * 6000#
    Else
        variables = 18# * re + 12# * edges + 3# * RR + 1#
        estimate = 6# * re * 1800# + variables * 2400# + edges * 12000#
    End If
    RPX_DiagEvent "assembly_estimate", "mode=" & mode & ";bytes=" & estimate & ";formula=conservative_v1_uncalibrated;includes_dictionary_allowance=true;not_a_guarantee=true"
    RPX_DiagEvent "p4_resident_cache", "bytes=" & RPX_SchurResidentBytes() & ";included_in_process_private_bytes=true"
    RPX_CapacityCheck "before_assembly_" & mode, estimate
End Sub

Public Function RPX_HoptMark() As Long
    If Not active Then Exit Function
    RPX_HoptMark = hoptSp
End Function
Public Sub RPX_HoptEnter(ByVal Bucket As Long)
    If Not active Then Exit Sub
    If Bucket < 1 Or Bucket > HoptBucketCount Then Exit Sub
    If hoptSp >= HoptStackMax Then
        hoptOverflow = hoptOverflow + 1
        Exit Sub
    End If
    hoptSp = hoptSp + 1
    hoptStack(hoptSp) = Bucket
    hoptStart(hoptSp) = ClockSeconds()
    hoptChild(hoptSp) = 0#
End Sub
Public Sub RPX_HoptLeave(ByVal Bucket As Long)
    Dim elapsed As Double
    If Not active Then Exit Sub
    If hoptSp <= 0 Then Exit Sub
    If hoptStack(hoptSp) <> Bucket Then Exit Sub
    elapsed = ClockSeconds() - hoptStart(hoptSp)
    hoptInclusive(Bucket) = hoptInclusive(Bucket) + elapsed
    hoptExclusive(Bucket) = hoptExclusive(Bucket) + (elapsed - hoptChild(hoptSp))
    hoptCalls(Bucket) = hoptCalls(Bucket) + 1
    hoptSp = hoptSp - 1
    If hoptSp > 0 Then hoptChild(hoptSp) = hoptChild(hoptSp) + elapsed
End Sub
Public Sub RPX_HoptUnwind(ByVal mark As Long)
    If Not active Then Exit Sub
    Do While hoptSp > mark
        RPX_HoptLeave hoptStack(hoptSp)
    Loop
End Sub
Public Sub RPX_HoptEpoch(ByVal iteration As Long, ByVal attempt As Long, ByVal matrixEpoch As Long, ByVal factorEpoch As Long, ByVal rhsEpoch As Long)
    If Not active Then Exit Sub
    hoptIteration = iteration
    hoptAttempt = attempt
    hoptMatrix = matrixEpoch
    hoptFactor = factorEpoch
    hoptRhs = rhsEpoch
End Sub
Public Sub RPX_HoptAddFinite(ByVal amount As Double)
    If Not active Then Exit Sub
    hoptFinite = hoptFinite + amount
    hoptSolveFinite = hoptSolveFinite + amount
End Sub
Public Sub RPX_HoptLuGauge(ByVal lNnz As Long, ByVal uNnz As Long, ByVal updateCount As Double, ByVal arena As Double, ByVal byteCount As Double, ByVal finiteCount As Double, Optional ByVal rowScanNodes As Double = 0, Optional ByVal rowScanRows As Double = 0, Optional ByVal indexDirect As Long = 0)
    If Not active Then Exit Sub
    hoptLuFactors = hoptLuFactors + 1
    hoptRowScanNodes = hoptRowScanNodes + rowScanNodes: hoptSolveRowScanNodes = hoptSolveRowScanNodes + rowScanNodes
    hoptRowScanRows = hoptRowScanRows + rowScanRows: hoptSolveRowScanRows = hoptSolveRowScanRows + rowScanRows
    hoptIndexDirect = hoptIndexDirect + indexDirect: hoptSolveIndexDirect = hoptSolveIndexDirect + indexDirect
    hoptUpdates = hoptUpdates + updateCount
    hoptSolveUpdates = hoptSolveUpdates + updateCount
    hoptFinite = hoptFinite + finiteCount
    hoptSolveFinite = hoptSolveFinite + finiteCount
    If arena > hoptMaxArena Then hoptMaxArena = arena
    If arena > hoptSolveMaxArena Then hoptSolveMaxArena = arena
    If byteCount > hoptMaxBytes Then hoptMaxBytes = byteCount
    If byteCount > hoptSolveMaxBytes Then hoptSolveMaxBytes = byteCount
    If lNnz > hoptMaxLnnz Then hoptMaxLnnz = lNnz
    If lNnz > hoptSolveMaxLnnz Then hoptSolveMaxLnnz = lNnz
    If uNnz > hoptMaxUnnz Then hoptMaxUnnz = uNnz
    If uNnz > hoptSolveMaxUnnz Then hoptSolveMaxUnnz = uNnz
End Sub
Private Sub HoptSnap()
    Dim i As Long
    For i = 1 To HoptBucketCount
        hoptMarkIncl(i) = hoptInclusive(i)
        hoptMarkExcl(i) = hoptExclusive(i)
        hoptMarkCalls(i) = hoptCalls(i)
    Next i
    hoptMarkFinite = hoptFinite
    hoptMarkUpdates = hoptUpdates
    hoptMarkFactors = hoptLuFactors
    hoptSolveFinite = 0#
    hoptSolveUpdates = 0#
    hoptSolveMaxArena = 0#
    hoptSolveMaxBytes = 0#
    hoptSolveMaxLnnz = 0
    hoptSolveMaxUnnz = 0
    hoptSolveRowScanNodes = 0#
    hoptSolveRowScanRows = 0#
    hoptSolveIndexDirect = 0
End Sub
Private Function HoptName(ByVal Bucket As Long) As String
    HoptName = Choose(Bucket, "ldl_numeric", "lu_store", "lu_eliminate", "lu_compress", "triangular_solve", "dd_operator", "dd_product", "border", "backward", "input_read", "input_compare", "live_key", "model_key", "stage_bar", "stage_file", "stage_log", "stage_yield", "input_assert")
End Function
Private Sub HoptEmit(ByVal scopeName As String)
    Dim i As Long, incl As Double, excl As Double, shown As Long, detail As String
    Dim finiteNow As Double, updatesNow As Double, factorsNow As Long
    If Not active Then Exit Sub
    For i = 1 To HoptBucketCount
        If scopeName = "solve" Then
            incl = hoptInclusive(i) - hoptMarkIncl(i)
            excl = hoptExclusive(i) - hoptMarkExcl(i)
            shown = hoptCalls(i) - hoptMarkCalls(i)
        Else
            incl = hoptInclusive(i)
            excl = hoptExclusive(i)
            shown = hoptCalls(i)
        End If
        If shown <> 0 Or incl <> 0# Or excl <> 0# Then
            detail = "scope=" & scopeName & ";bucket=" & HoptName(i) & ";inclusive_seconds=" & CStr(incl) & ";exclusive_seconds=" & CStr(excl) & ";calls=" & CStr(shown)
            RPX_DiagEvent "hopt0", detail
        End If
    Next i
    If scopeName = "solve" Then
        finiteNow = hoptSolveFinite
        updatesNow = hoptSolveUpdates
        factorsNow = hoptLuFactors - hoptMarkFactors
        detail = "scope=solve;lu_factors=" & factorsNow & ";updates=" & CStr(updatesNow) & ";finite_checks=" & CStr(finiteNow)
        detail = detail & ";max_arena=" & CStr(hoptSolveMaxArena) & ";max_bytes=" & CStr(hoptSolveMaxBytes)
        detail = detail & ";max_lnnz=" & hoptSolveMaxLnnz & ";max_unnz=" & hoptSolveMaxUnnz
        detail = detail & ";row_scan_nodes=" & CStr(hoptSolveRowScanNodes) & ";row_scan_rows=" & CStr(hoptSolveRowScanRows) & ";index_direct=" & CStr(hoptSolveIndexDirect)
    Else
        detail = "scope=run;lu_factors=" & hoptLuFactors & ";updates=" & CStr(hoptUpdates) & ";finite_checks=" & CStr(hoptFinite)
        detail = detail & ";max_arena=" & CStr(hoptMaxArena) & ";max_bytes=" & CStr(hoptMaxBytes)
        detail = detail & ";max_lnnz=" & hoptMaxLnnz & ";max_unnz=" & hoptMaxUnnz
        detail = detail & ";row_scan_nodes=" & CStr(hoptRowScanNodes) & ";row_scan_rows=" & CStr(hoptRowScanRows) & ";index_direct=" & CStr(hoptIndexDirect)
    End If
    detail = detail & ";iteration=" & hoptIteration & ";attempt=" & hoptAttempt & ";matrix_epoch=" & hoptMatrix & ";factor_epoch=" & hoptFactor & ";rhs_epoch=" & hoptRhs
    detail = detail & ";open_frames=" & hoptSp & ";overflow=" & hoptOverflow
    RPX_DiagEvent "hopt0_gauge", detail
End Sub




Public Sub RPX_DiagAuditSave(ByRef state As RPX_AuditSnapshot)
    Dim i As Long
    state.mode = auditMode
    For i = 1 To 13
        state.maximum(i) = auditMaximum(i): state.location(i) = auditLocation(i): state.seen(i) = auditSeen(i)
    Next i
End Sub
Public Sub RPX_DiagAuditRestore(ByRef state As RPX_AuditSnapshot)
    Dim i As Long
    auditMode = state.mode
    For i = 1 To 13
        auditMaximum(i) = state.maximum(i): auditLocation(i) = state.location(i): auditSeen(i) = state.seen(i)
    Next i
End Sub
Public Function RPX_DiagAuditSummary() As String
    Dim i As Long
    For i = 1 To 13
        If auditSeen(i) Then RPX_DiagAuditSummary = RPX_DiagAuditSummary & AuditName(i) & "=" & auditMaximum(i) & ";" & auditLocation(i) & ";"
    Next i
End Function
