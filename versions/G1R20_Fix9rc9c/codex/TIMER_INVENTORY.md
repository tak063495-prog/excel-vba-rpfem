# Timer inventory - rc9b reference and rc9c disposition

All line numbers below refer to untouched `vba/` CP932 sources, unless expressly labeled rc9c. Every file, including `ThisWorkbook.code.txt`, was scanned case-insensitively for word-boundary `Timer`, `Now`, `Time`, `Date`, and timer-wrapper names. No `Time()` elapsed use was found. No Rnd seed uses Timer in this tree.

## Every Timer occurrence: 40 lines in 15 modules

| Module | Procedure | Line | Purpose | Elapsed / deadline classification | rc9c action |
|---|---|---:|---|---|---|
| RPX_AccuracyNative.bas | AccuracySolveCore | 356 | DLL poll / watchdog | capture; later elapsed | RPX_Clock capture |
| RPX_AccuracyNative.bas | AccuracySolveCore | 368 | DLL poll / watchdog | elapsed vs seconds + 5 deadline; single-wrap corrected before | RPX_Clock difference |
| RPX_AccuracyNative.bas | DirectEqualityRepair | 708 | log: direct equality repair | capture; later elapsed | RPX_Clock capture |
| RPX_AccuracyNative.bas | DirectEqualityRepair | 767 | log: direct equality repair | elapsed difference | RPX_Clock difference |
| RPX_Adapt.bas | Fix8ChainStep | 3413 | log: thin round step | capture; later elapsed | RPX_Clock capture |
| RPX_Adapt.bas | Fix8ChainStep | 3514 | log: thin round step | elapsed difference; P06N defect | RPX_Clock difference |
| RPX_AHO3.bas | RPX_AHOCheckBudget | 35 | time protection: AHO global budget | elapsed vs budget; single-wrap corrected before | RPX_Clock difference |
| RPX_AHO3.bas | RPX_AHOFixedWitness | 321 | time protection: AHO global budget | capture; later elapsed | RPX_Clock capture |
| RPX_AHO3.bas | RPX_AHOFixedWitness | 339 | time protection: AHO iteration budget | capture; later elapsed | RPX_Clock capture |
| RPX_AHO3.bas | RPX_AHOFixedWitness | 369 | time protection: AHO iteration budget | elapsed vs budget; single-wrap corrected before | RPX_Clock difference |
| RPX_Conic.bas | RPX_ContextGuardDue | 140 | other: context guard cadence | capture; elapsed at 142 (corrected 143) | RPX_Clock capture |
| RPX_Conic.bas | RPX_Stage | 163 | other: sparse stage cadence | capture; elapsed at 165 (corrected 166) | RPX_Clock capture |
| RPX_Conic.bas | RPX_Stage | 192 | log: raw local seconds-of-day | neither elapsed nor deadline | retained Timer |
| RPX_Control.bas | RPX_Run | 180 | log: worksheet analysis duration | capture; later elapsed | RPX_Clock capture |
| RPX_Control.bas | RPX_Run | 267 | log: worksheet analysis duration | elapsed modulo one day | inactive original; ASCII live monotonic difference without Mod |
| RPX_Diagnostics.bas | ClockSeconds | 173 | other: diagnostic clock fallback for all elapsed users | absolute full-date seconds; indirect elapsed/deadline | ClockSeconds delegates to RPX_Clock |
| RPX_Fs.bas | RPX_TestFs | 1017 | log: Fs test output | capture; later elapsed | RPX_Clock capture |
| RPX_Fs.bas | RPX_TestFs | 1020 | log: Fs test output | elapsed difference | RPX_Clock difference |
| RPX_Lower.bas | RPX_TestModel | 257 | log: model test output | capture; later elapsed | RPX_Clock capture |
| RPX_Lower.bas | RPX_TestModel | 263 | log: model test output | elapsed difference | RPX_Clock difference |
| RPX_Lower.bas | RPX_TestOrderModel | 273 | log: order test output | capture; later elapsed | RPX_Clock capture |
| RPX_Lower.bas | RPX_TestOrderModel | 277 | log: order test output | elapsed difference | RPX_Clock difference |
| RPX_LowerCache.bas | SecondsSince | 18 | log: lower-cache SecondsSince helper | elapsed difference; corrected at 19 | RPX_Clock difference |
| RPX_LowerCache.bas | RPX_LowerCacheMatch | 93 | log: cache match timing | capture; SecondsSince consumer | RPX_Clock capture |
| RPX_LowerCache.bas | RPX_LowerCacheStore | 114 | log: cache store timing | capture; SecondsSince consumer | RPX_Clock capture |
| RPX_LowerCache.bas | RPX_LowerCacheRestore | 172 | log: cache restore timing | capture; SecondsSince consumer | RPX_Clock capture |
| RPX_NativeBulk.bas | RPX_NativePulse | 116 | DLL pulse: UI and input watchdog cadence | capture; elapsed at 120/127 (both corrected before) | RPX_Clock capture |
| RPX_Schur.bas | RPX_SchurFactor | 416 | other: Schur progress pulse cadence | lastPulse capture; later elapsed | RPX_Clock capture |
| RPX_Schur.bas | CompareCore | 609 | log: KKT comparison time | capture; later elapsed | RPX_Clock capture |
| RPX_Schur.bas | CompareCore | 659 | log: KKT comparison time | elapsed difference | RPX_Clock difference |
| RPX_Schur.bas | Pulse | 740 | other: Schur progress cadence | elapsed at 741; backward-clock branch before | RPX_Clock read; now renamed nowClock |
| RPX_Schur.bas | Pulse | 743 | other: Schur progress cadence | lastPulse capture; later elapsed | RPX_Clock capture |
| RPX_SeepUI.bas | RPX_SeepRun | 1135 | log: seep status duration | capture; later elapsed | RPX_Clock capture |
| RPX_SeepUI.bas | RPX_SeepRun | 1216 | log: seep status duration | elapsed difference | inactive original; ASCII live monotonic difference |
| RPX_SeepWT.bas | RPX_SeepWaterTableSelfTest | 431 | log: water-table self-test duration | capture; later elapsed | RPX_Clock capture |
| RPX_SeepWT.bas | RPX_SeepWaterTableSelfTest | 466 | log: water-table self-test duration | elapsed difference | inactive original; ASCII live monotonic difference |
| RPX_Seismic.bas | OneCase | 171 | log: seismic case duration | capture; later elapsed | RPX_Clock capture |
| RPX_Seismic.bas | OneCase | 176 | log: seismic case duration | elapsed difference | inactive original; ASCII live monotonic difference |
| RPX_Solve.bas | RPX_TestConic | 2405 | log: conic test duration | capture; later elapsed | RPX_Clock capture |
| RPX_Solve.bas | RPX_TestConic | 2408 | log: conic test duration | elapsed difference | RPX_Clock difference |

## Every Now / Time / Date word match (including non-code matches)

The five elapsed `Now` occurrences are one start capture and four later reads of the same clock. Full-date `Now` was already safe across an ordinary midnight; it is migrated to remove wall-clock adjustment sensitivity and to exercise the same test hook. Other `Now` calls remain calendar timestamps. The Schur local named `now` is a variable, and is renamed during this patch.

| Module | Procedure | Line | Purpose / elapsed or deadline? | Original line |
|---|---|---:|---|---|
| RPX_Adapt.bas | OptionalStageExpired | 188 | time protection: optional-stage elapsed difference vs deadline; migrated to RPX_Clock | `elapsed = CDbl(Now) * 86400# - adaptStartedAt` |
| RPX_Adapt.bas | (module) | 1318 | comment only; no clock operation | `' already go straight to HSD and the last upper HSD time scaled linearly to the` |
| RPX_Adapt.bas | (module) | 1320 | comment only; no clock operation | `' end in HSD_SOLVE_BUDGET_EXCEEDED; stop it before spending that time. The kept` |
| RPX_Adapt.bas | RunAdaptCore | 1956 | time protection: adaptation start capture for later elapsed/deadline reads; migrated to RPX_Clock | `adaptStartedAt = CDbl(Now) * 86400#` |
| RPX_Adapt.bas | Fix9ThinEarlyBudget | 2803 | time protection: early-admission elapsed difference, feeding remaining allowance; migrated to RPX_Clock | `elapsed = CDbl(Now) * 86400# - adaptStartedAt` |
| RPX_Adapt.bas | Fix9DeadlineRemainingSeconds | 2954 | time protection: case-limit fallback elapsed difference; migrated to RPX_Clock | `If caseElapsedSeconds < 0# And adaptStartedAt > 0# Then caseElapsedSeconds = CDbl(Now) * 86400# - adaptStartedAt` |
| RPX_Adapt.bas | Fix9DeadlineRemainingSeconds | 2973 | time protection: deadline/nominal remaining = limit minus elapsed difference; migrated to RPX_Clock | `remaining = limit - (CDbl(Now) * 86400# - adaptStartedAt)` |
| RPX_Adapt.bas | Fix9AuxUpper | 3102 | comment only; no clock operation | `' is accepted unless ADAPT_THIN_AUX_ALMOST=0. Time-limit/error iterates are not.` |
| RPX_C0Guides.bas | (module) | 6 | comment only; no clock operation | `' The automatic Fs path (RPX_C0Select) turns guides on for one candidate at a time.` |
| RPX_Diagnostics.bas | ClockSeconds | 173 | clock helper fallback: full-date seconds for elapsed users; replaced with RPX_Clock | `ClockSeconds = CDbl(Date) * 86400# + Timer` |
| RPX_Diagnostics.bas | RPX_DiagEvent | 243 | calendar timestamp / file or result identifier; no elapsed difference or deadline; retained | `line = Format$(now, "yyyy-mm-dd hh:nn:ss") & vbTab & Format$(ClockSeconds() - started, "0.000000") & vbTab & solveId & vbTab & Clean(candidate) & vbTab & solveMode & vbTab & CStr(solveFactor) & vbTab & eventName & vbTab & mem & vbTab & Clean(detail)` |
| RPX_Diagnostics.bas | RPX_DiagStart | 279 | calendar timestamp / file or result identifier; no elapsed difference or deadline; retained | `base = folder & Application.PathSeparator & Format$(now, "yyyymmdd_hhnnss") & "_" & CStr(RPX_GetPID()) & "_" & operation` |
| RPX_Robust.bas | (module) | 10 | comment only; no clock operation | `' this input/policy (P05A: conic succeeds at most Fs points; skipping cost time).` |
| RPX_Schur.bas | Pulse | 738 | local variable named now, not builtin Now; cadence elapsed; renamed nowClock | `Dim now As Double` |
| RPX_Schur.bas | Pulse | 740 | local variable named now, not builtin Now; cadence elapsed; renamed nowClock | `now = Timer` |
| RPX_Schur.bas | Pulse | 741 | local variable named now, not builtin Now; cadence elapsed; renamed nowClock | `If now < lastPulse Or now - lastPulse >= 0.2 Then` |
| RPX_Seep.bas | NextSeepResultId | 343 | calendar timestamp / file or result identifier; no elapsed difference or deadline; retained | `NextSeepResultId = Format(now, "yyyymmddhhnnss") & "-" & SN & "-" & SM & "-" & CStr(seepWriteSequence)` |
| RPX_SeepUI.bas | RPX_SeepTrace | 105 | calendar timestamp / file or result identifier; no elapsed difference or deadline; retained | `Print #f, Format$(now, "yyyy-mm-dd hh:nn:ss") & " " & text` |
| RPX_Solve.bas | HNativeCapture | 914 | calendar timestamp / file or result identifier; no elapsed difference or deadline; retained | `path = ThisWorkbook.path & "\hsd_linear_" & Format$(now, "yyyymmdd_hhnnss") & "_" & P4Solve & "_" & XIteration & "_" & hCaptureCount & ".tsv"` |
| RPX_Solve.bas | (module) | 1962 | comment only; no clock operation | `' G1R20 Fix6 (#29): HSD time limit with an endgame grace. When the limit is` |
| RPX_Solve.bas | BudgetCheck | 2007 | comment only; no clock operation | `' Thin case/protected time is a hard wall, including any endgame grace.` |
| RPX_Solve.bas | (module) | 2024 | comment only; no clock operation | `' G1R20 Fix7 (candidate): the ordinary HSD time limit for a mesh of the given` |
| RPX_Solve.bas | OptimizeHSDCore | 2245 | comment only; no clock operation | `' G1R20 Fix6 (#29, P10A): the time limit is checked only AFTER the current` |

## Timer-derived helper definitions and every named occurrence: 74 lines

`ClockSeconds` originally selected QPC or `Date*86400+Timer`; `RPX_DiagClock` exposed it. `SecondsSince` directly wrapped Timer; `P4Time` and `RPX_ThinRunElapsedSeconds` wrapped `RPX_DiagClock`. In rc9c these routes all reach the single public `RPX_Clock` seconds clock. This table includes declarations and call sites so hidden elapsed consumers and synthesized DLL start captures are reviewable.

| Module | Procedure | Line | Purpose | Elapsed / deadline classification | Original line |
|---|---|---:|---|---|---|
| RPX_AccuracyNative.bas | ThinUpperStatus5Retry | 448 | DLL/recovery time budget or retry duration | capture/current clock read; feeds elapsed timing | `retryStarted = RPX_DiagClock()` |
| RPX_AccuracyNative.bas | ThinUpperStatus5Retry | 461 | DLL/recovery time budget or retry duration | elapsed difference / budget or duration arithmetic | `RPX_DiagEvent "thin_layer_dll_retry", "mode=upper;Fs=" & pr.factor & ";initial_status=5;initial_iterations=" & initialIterations & ";attempts=1;iteration_limit=" & RPX_ThinUpperDllRetryIterations & ";equilibrate=" & retryEquilibratePacked & ";seconds_limit=" & retrySeconds & ";seconds=" & RPX_DiagClock() - retryStarted & ";solver_status=" & retryStatus & ";iterations=" & retryIterations & ";finite_candidate=" & retryFinite & ";original_audit_required=1;bound=0"` |
| RPX_AccuracyNative.bas | RPX_AccuracyDirectBound | 1062 | DLL/recovery time budget or retry duration | capture/current clock read; feeds elapsed timing | `started = RPX_DiagClock()` |
| RPX_AccuracyNative.bas | RPX_AccuracyDirectBound | 1118 | DLL/recovery time budget or retry duration | elapsed difference / budget or duration arithmetic | `remaining = seconds - (RPX_DiagClock() - started)` |
| RPX_AccuracyNative.bas | TryOriginalDualLinearRepair | 1536 | DLL/recovery time budget or retry duration | capture/current clock read; feeds elapsed timing | `started = RPX_DiagClock()` |
| RPX_AccuracyNative.bas | TryOriginalDualLinearRepair | 1538 | DLL/recovery time budget or retry duration | elapsed difference / budget or duration arithmetic | `seconds = RPX_DiagClock() - started` |
| RPX_Adapt.bas | Fix9DeadlineRemainingSeconds | 2953 | round time / time protection / recovery allowance | capture/current clock read; feeds elapsed timing | `caseElapsedSeconds = RPX_ThinRunElapsedSeconds()` |
| RPX_Adapt.bas | Fix9UpperProtectBegin | 2982 | round time / time protection / recovery allowance | capture/current clock read; feeds elapsed timing | `fix9UpperProtectStarted = RPX_DiagClock()` |
| RPX_Adapt.bas | RPX_ThinUpperProtectLimit | 3018 | round time / time protection / recovery allowance | elapsed difference / budget or duration arithmetic | `elapsedSeconds = RPX_DiagClock() - fix9UpperProtectStarted` |
| RPX_Adapt.bas | Fix9AuxUpper | 3097 | round time / time protection / recovery allowance | capture/current clock read; feeds elapsed timing | `started = RPX_DiagClock(): reason = "solve_failed_or_timeout"` |
| RPX_Adapt.bas | Fix9AuxUpper | 3099 | round time / time protection / recovery allowance | elapsed difference / budget or duration arithmetic | `If RPX_DiagClock() - started > seconds Then reason = "time_limit": GoTo AuxDone` |
| RPX_Adapt.bas | Fix8ThinLayerChains | 3812 | round time / time protection / recovery allowance | capture/current clock read; feeds elapsed timing | `lowerStarted = RPX_DiagClock()` |
| RPX_Adapt.bas | Fix8ThinLayerChains | 3820 | round time / time protection / recovery allowance | elapsed difference / budget or duration arithmetic | `opt.lowerCostSeconds = RPX_DiagClock() - lowerStarted: opt.lowerCostElements = loLine.elements` |
| RPX_Diagnostics.bas | RPX_DiagClock | 165 | other: profiling / log | helper declaration | `Public Function RPX_DiagClock() As Double` |
| RPX_Diagnostics.bas | RPX_DiagClock | 166 | other: profiling / log | capture/current clock read; feeds elapsed timing | `RPX_DiagClock = ClockSeconds()` |
| RPX_Diagnostics.bas | ClockSeconds | 168 | other: profiling / log | helper declaration | `Private Function ClockSeconds() As Double` |
| RPX_Diagnostics.bas | ClockSeconds | 171 | other: profiling / log | helper declaration | `If RPX_QPC(count) <> 0 Then ClockSeconds = count / frequency: Exit Function` |
| RPX_Diagnostics.bas | ClockSeconds | 173 | other: profiling / log | capture/current clock read; feeds elapsed timing | `ClockSeconds = CDbl(Date) * 86400# + Timer` |
| RPX_Diagnostics.bas | FlushLog | 219 | other: profiling / log | capture/current clock read; feeds elapsed timing | `logStarted = ClockSeconds()` |
| RPX_Diagnostics.bas | FlushLog | 225 | other: profiling / log | elapsed difference / budget or duration arithmetic | `loggingSeconds = loggingSeconds + ClockSeconds() - logStarted` |
| RPX_Diagnostics.bas | FlushLog | 232 | other: profiling / log | elapsed difference / budget or duration arithmetic | `loggingSeconds = loggingSeconds + ClockSeconds() - logStarted` |
| RPX_Diagnostics.bas | RPX_DiagEvent | 243 | other: profiling / log | elapsed difference / budget or duration arithmetic | `line = Format$(now, "yyyy-mm-dd hh:nn:ss") & vbTab & Format$(ClockSeconds() - started, "0.000000") & vbTab & solveId & vbTab & Clean(candidate) & vbTab & solveMode & vbTab & CStr(solveFactor) & vbTab & eventName & vbTab & mem & vbTab & Clean(detail)` |
| RPX_Diagnostics.bas | RPX_DiagStart | 267 | other: profiling / log | capture/current clock read; feeds elapsed timing | `started = ClockSeconds(): active = True` |
| RPX_Diagnostics.bas | RPX_DiagBegin | 307 | other: profiling / log | capture/current clock read; feeds elapsed timing | `phaseStart(phase) = ClockSeconds()` |
| RPX_Diagnostics.bas | RPX_DiagEnd | 316 | other: profiling / log | elapsed difference / budget or duration arithmetic | `elapsed = ClockSeconds() - phaseStart(phase)` |
| RPX_Diagnostics.bas | RPX_DiagSolveStart | 341 | other: profiling / log | capture/current clock read; feeds elapsed timing | `solveId = solveId + 1: solveMode = mode: solveFactor = factor: solveStart = ClockSeconds(): inSolve = True` |
| RPX_Diagnostics.bas | RPX_DiagSolveEnd | 367 | other: profiling / log | elapsed difference / budget or duration arithmetic | `RPX_DiagEvent "solve_end", "status=" & state & ";wall_seconds=" & CStr(ClockSeconds() - solveStart) & ";last_iteration_index=" & XIteration & ";objective=" & XObjective & ";primal=" & XPrimalResidual & ";dual=" & XDualResidual & ";gap=" & XGap & ";message=" & message` |
| RPX_Diagnostics.bas | RPX_HoptEnter | 480 | other: profiling / log | capture/current clock read; feeds elapsed timing | `hoptStart(hoptSp) = ClockSeconds()` |
| RPX_Diagnostics.bas | RPX_HoptLeave | 488 | other: profiling / log | elapsed difference / budget or duration arithmetic | `elapsed = ClockSeconds() - hoptStart(hoptSp)` |
| RPX_Diagnostics.bas | RPX_ThinRunElapsedSeconds | 620 | time protection: full run elapsed for case limit | helper declaration | `Public Function RPX_ThinRunElapsedSeconds() As Double` |
| RPX_Diagnostics.bas | RPX_ThinRunElapsedSeconds | 621 | time protection: full run elapsed for case limit | capture/current clock read; feeds elapsed timing | `RPX_ThinRunElapsedSeconds = -1#` |
| RPX_Diagnostics.bas | RPX_ThinRunElapsedSeconds | 623 | time protection: full run elapsed for case limit | elapsed difference / budget or duration arithmetic | `RPX_ThinRunElapsedSeconds = RPX_DiagClock() - started` |
| RPX_Guard.bas | RPX_BenchInputCompare | 332 | other: profiling / log | capture/current clock read; feeds elapsed timing | `t0 = RPX_DiagClock()` |
| RPX_Guard.bas | RPX_BenchInputCompare | 336 | other: profiling / log | elapsed difference / budget or duration arithmetic | `oldSeconds = oldSeconds + (RPX_DiagClock() - t0)` |
| RPX_Guard.bas | RPX_BenchInputCompare | 337 | other: profiling / log | capture/current clock read; feeds elapsed timing | `t0 = RPX_DiagClock()` |
| RPX_Guard.bas | RPX_BenchInputCompare | 341 | other: profiling / log | elapsed difference / budget or duration arithmetic | `newSeconds = newSeconds + (RPX_DiagClock() - t0)` |
| RPX_LowerCache.bas | SecondsSince | 17 | other: profiling / log | helper declaration | `Private Function SecondsSince(ByVal started As Double) As Double` |
| RPX_LowerCache.bas | SecondsSince | 18 | other: profiling / log | elapsed difference / budget or duration arithmetic | `SecondsSince = Timer - started` |
| RPX_LowerCache.bas | SecondsSince | 19 | other: profiling / log | capture/current clock read; feeds elapsed timing | `If SecondsSince < 0# Then SecondsSince = SecondsSince + 86400#` |
| RPX_LowerCache.bas | RPX_LowerCacheMatch | 108 | other: profiling / log | capture/current clock read; feeds elapsed timing | `RPX_DiagEvent "lower_cache_lookup", "hit=" & Abs(CLng(RPX_LowerCacheMatch)) & ";reason=" & reason & ";seconds=" & SecondsSince(started)` |
| RPX_LowerCache.bas | RPX_LowerCacheStore | 124 | other: profiling / log | capture/current clock read; feeds elapsed timing | `RPX_DiagEvent "lower_cache_keep", "bytes=" & cacheBytes & ";seconds=" & SecondsSince(started) & ";original_q=" & savedOriginal.subdivisions` |
| RPX_LowerCache.bas | RPX_LowerCacheStore | 167 | other: profiling / log | capture/current clock read; feeds elapsed timing | `RPX_DiagEvent "lower_cache_store", "engine=20260927_p3b;bytes=" & bytes & ";seconds=" & SecondsSince(started) & ";xn=" & XN & ";xm=" & XM & ";xb=" & XB` |
| RPX_LowerCache.bas | RPX_LowerCacheRestore | 202 | other: profiling / log | capture/current clock read; feeds elapsed timing | `RPX_DiagEvent "lower_cache_restore", "bytes=" & cacheBytes & ";seconds=" & SecondsSince(started) & ";prepared=1;cold_start=1;symbolic=rebuilt;original_q=" & XOriginalCompiled.subdivisions` |
| RPX_P4.bas | RPX_P4Time | 231 | other: profiling / log | elapsed difference / budget or duration arithmetic | `phaseTime(phase) = phaseTime(phase) + RPX_DiagClock() - started` |
| RPX_Robust.bas | RPX_OptimizeBound | 116 | round time: last upper HSD prediction | capture/current clock read; feeds elapsed timing | `rescueStarted = RPX_DiagClock()` |
| RPX_Robust.bas | RPX_OptimizeBound | 118 | round time: last upper HSD prediction | elapsed difference / budget or duration arithmetic | `If mode = "upper" Then RPX_LastUpperHsdSeconds = RPX_DiagClock() - rescueStarted: RPX_LastUpperHsdElements = re` |
| RPX_Schur.bas | NativeFactor | 184 | other: profiling / log | synthesized capture from DLL milliseconds for P4Time | `RPX_P4Time 1, RPX_DiagClock() - metrics(1)` |
| RPX_Schur.bas | NativeFactor | 185 | other: profiling / log | synthesized capture from DLL milliseconds for P4Time | `RPX_P4Time 2, RPX_DiagClock() - metrics(2)` |
| RPX_Schur.bas | NativeSolve | 220 | other: profiling / log | synthesized capture from DLL milliseconds for P4Time | `RPX_P4Time 4, RPX_DiagClock() - metrics(8)` |
| RPX_Schur.bas | NativeSolve | 221 | other: profiling / log | synthesized capture from DLL milliseconds for P4Time | `RPX_P4Time 5, RPX_DiagClock() - metrics(9)` |
| RPX_Schur.bas | NativeSolve | 222 | other: profiling / log | synthesized capture from DLL milliseconds for P4Time | `RPX_P4Time 6, RPX_DiagClock() - metrics(10)` |
| RPX_Schur.bas | NativeSolve | 223 | other: profiling / log | synthesized capture from DLL milliseconds for P4Time | `RPX_P4Time 7, RPX_DiagClock() - metrics(11)` |
| RPX_Schur.bas | RPX_SchurFactor | 416 | other: profiling / log | capture/current clock read; feeds elapsed timing | `started = RPX_DiagClock(): lastPulse = Timer` |
| RPX_Schur.bas | RPX_SchurFactor | 418 | other: profiling / log | capture/current clock read; feeds elapsed timing | `phaseStart = RPX_DiagClock()` |
| RPX_Schur.bas | RPX_SchurFactor | 447 | other: profiling / log | capture/current clock read; feeds elapsed timing | `phaseStart = RPX_DiagClock()` |
| RPX_Schur.bas | RPX_SchurFactor | 460 | other: profiling / log | capture/current clock read; feeds elapsed timing | `phaseStart = RPX_DiagClock()` |
| RPX_Schur.bas | RPX_SchurFactor | 496 | other: profiling / log | capture/current clock read; feeds elapsed timing | `RPX_P4Time 1, phaseStart: phaseStart = RPX_DiagClock()` |
| RPX_Schur.bas | RPX_SchurFactor | 509 | other: profiling / log | capture/current clock read; feeds elapsed timing | `RPX_P4Time 2, phaseStart: phaseStart = RPX_DiagClock()` |
| RPX_Schur.bas | RPX_SchurFactor | 511 | other: profiling / log | capture/current clock read; feeds elapsed timing | `phaseStart = RPX_DiagClock()` |
| RPX_Schur.bas | RPX_SchurFactor | 520 | other: profiling / log | elapsed difference / budget or duration arithmetic | `RPX_DiagEvent "schur_numeric", "iteration=" & XIteration & ";delta=" & reg & ";seconds=" & RPX_DiagClock() - started` |
| RPX_Schur.bas | SolveReduced | 536 | other: profiling / log | capture/current clock read; feeds elapsed timing | `P4Reduced = P4Reduced + 1: started = RPX_DiagClock()` |
| RPX_Schur.bas | SolveReduced | 554 | other: profiling / log | capture/current clock read; feeds elapsed timing | `RPX_P4Time 4, started: started = RPX_DiagClock()` |
| RPX_Schur.bas | SolveReduced | 556 | other: profiling / log | capture/current clock read; feeds elapsed timing | `RPX_P4Time 5, started: started = RPX_DiagClock()` |
| RPX_Schur.bas | RPX_SchurAutoSolve | 690 | other: profiling / log | capture/current clock read; feeds elapsed timing | `correctionStart = RPX_DiagClock(): started = correctionStart` |
| RPX_Schur.bas | RPX_SchurAutoSolve | 708 | other: profiling / log | capture/current clock read; feeds elapsed timing | `started = RPX_DiagClock()` |
| RPX_Solve.bas | EnsureGeneric | 471 | other: profiling / log | capture/current clock read; feeds elapsed timing | `started = RPX_DiagClock()` |
| RPX_Solve.bas | OptimizeCore | 889 | other: profiling / log | capture/current clock read; feeds elapsed timing | `replayed = True: replayStart = RPX_DiagClock(): P4Replays = P4Replays + 1` |
| RPX_Solve.bas | ThinHsdBudgetClamp | 1979 | HSD budget / hard wall | elapsed difference / budget or duration arithmetic | `wallElapsed = RPX_DiagClock() - solveStarted` |
| RPX_Solve.bas | ThinHsdWallCheck | 1992 | HSD budget / hard wall | elapsed difference / budget or duration arithmetic | `elapsed = RPX_DiagClock() - solveStarted` |
| RPX_Solve.bas | BudgetCheck | 2006 | HSD budget / hard wall | elapsed difference / budget or duration arithmetic | `elapsed = RPX_DiagClock() - solveStarted` |
| RPX_Solve.bas | OptimizeHSDCore | 2101 | HSD budget / hard wall | capture/current clock read; feeds elapsed timing | `solveStarted = RPX_DiagClock()` |
| RPX_Solve.bas | OptimizeHSDCore | 2142 | HSD budget / hard wall | elapsed difference / budget or duration arithmetic | `thinHardWallSeconds = RPX_DiagClock() - solveStarted + upperAvailable` |
| RPX_Solve.bas | OptimizeHSDCore | 2165 | HSD budget / hard wall | elapsed difference / budget or duration arithmetic | `thinHardWallSeconds = RPX_DiagClock() - solveStarted + lowerAvailable` |
| RPX_Solve.bas | OptimizeHSDCore | 2176 | HSD budget / hard wall | elapsed difference / budget or duration arithmetic | `protectedWallSeconds = RPX_DiagClock() - solveStarted + protectedWallSeconds` |

## rc9c remaining Timer and test-hook reads

Three executable raw Timer reads remain inside the new `RPX_Clock.bas` (imported module name `RPX_ClockModule`): normal/fault fallback raw input and the arming/disarming alignment samples. They do not form an uncorrected elapsed difference. The one other executable Timer read is the original Conic stage time-of-day log. No remaining elapsed Timer use exists in active code.

- `RPX_Clock.bas:25 rawSeconds = CDbl(Timer)`
- `RPX_Clock.bas:64 rawSeconds = CDbl(Timer)`
- `RPX_Clock.bas:76 rawSeconds = CDbl(Timer)`
- `RPX_Conic.bas:190 f = FreeFile: Open XLogPath For Append As #f: Print #f, "# " & message & " " & CStr(Timer): Close #f`

Four original Japanese lines are retained byte-identically inside `#If False Then` branches, with ASCII live equivalents in `#Else`. Raw grep therefore still finds inactive Timer expressions. Their rc9c locations are: RPX_Control.bas:268, RPX_SeepUI.bas:1217, RPX_SeepWT.bas:467, RPX_Seismic.bas:177. They do not compile or run. All other non-ASCII lines remain byte-identical and in their original order.

The new helper contains no `Now`, `Time`, or `Date` reads. Default Win64/VBA7 mode calls GetTickCount64 and converts milliseconds to Double seconds. Other modes unwrap Timer by adding 86400 seconds whenever a sampled raw value decreases. With `RPX_TestClockWrap(120)`, the same fallback receives a rotated raw seconds-of-day stream that drops by 86400 around 120 seconds after arming. Test configuration is idle-only and gated by TEST_INJECTION=1; all module-level test state defaults to zero/off.

All elapsed start captures, saved pulse clocks, and differences were checked as Double. Remaining limits, reserves, nominal priority, iteration checks, solver settings and certificates are unchanged. The fallback needs at least one read per day to observe each wrap and cannot distinguish manual backward wall-clock changes from midnight; production Win64 uses the uptime API.

An additional scan for DateDiff, DateAdd, TimeValue, TimeSerial, OnTime, Application.Wait, Sleep, GetTickCount, timeGetTime and performance-counter APIs found only the already-audited QPC/QPF declarations (`RPX_Diagnostics:50-51`) and `Idle`, a Sleep alias (`RPX_AccuracyNative:14`) used between poll iterations. Sleep is a wait, not an elapsed/deadline clock, and is unchanged. No additional elapsed clock was found.
