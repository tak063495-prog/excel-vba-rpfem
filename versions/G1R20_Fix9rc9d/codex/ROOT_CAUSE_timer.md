# rc9c Timer root cause and P06N impact

## Conclusion

The supplied P06N log proves an incorrect negative thin-step duration across midnight. It does **not** show that the thin remaining-time calculation or HSD hard wall was disabled by that crossing. Those protections used a different clock in rc9b, and the post-midnight HSD budget expiry and refused solver starts demonstrate continued enforcement.

All code references in the P06N analysis below refer to the original, unchanged `vba/` sources. TSV row numbers count the header as row 1. The supplied log is an excerpt beginning at elapsed 15010.274235 seconds, so it does not provide the run's initial settings events or a complete history.

## Root cause

VBA `Timer` is seconds since local midnight. For start `s` before midnight and end `e` after midnight, an uncorrected `e - s` is the real duration minus 86400 seconds. The result remains negative until the local time of day catches up with the start time.

`RPX_Adapt.bas:3413` captures `started = Timer` in `Fix8ChainStep`. Its later read at `RPX_Adapt.bas:3514` computes `Timer - started` solely for the `thin_layer_step` diagnostic event. The P06N upper round 8 started around 23:31:07 and finished at 00:10:21. Log row 222 records `seconds=-84046.3`; adding 86400 gives 2353.7 seconds, approximately 39 minutes 14 seconds, consistent with the surrounding timestamps and diagnostic elapsed values.

That local `started` variable does not feed thin remaining time, solver budgets, measured lower-round cost, acceptance, or the certified bounds. A negative report therefore cannot by itself establish a failed budget or explain the unmet P06N gap target.

Uncorrected elapsed reporting also existed in several command/test/UI paths. Other Timer paths already corrected one midnight crossing. The inventory distinguishes all of these from absolute timestamp uses.

## The rc9b clocks actually used by time protection

### Thin remaining time and setting priority

`Fix9DeadlineRemainingSeconds` (`RPX_Adapt.bas:2947-2976`) returns a clamped, nonnegative remaining duration when a clock is available, or `-1` when unavailable. Its setting priority is:

1. Positive `ADAPT_THIN_CASE_LIMIT_SECONDS` (`:2951-2961`). It subtracts diagnostic run elapsed from the case limit. `RPX_ThinRunElapsedSeconds` (`RPX_Diagnostics.bas:620-624`) reads `RPX_DiagClock() - started`; if unavailable, the adaptation elapsed clock is the fallback (`RPX_Adapt.bas:2954`). The case limit applies to elapsed run time, which can begin before adaptation.
2. Positive `ADAPT_OPTIONAL_DEADLINE_SECONDS` (`RPX_Adapt.bas:2963-2967`).
3. Otherwise positive nominal thin allowance (`:2968-2969`), read from `ADAPT_THIN_NOMINAL_SECONDS`, default 18171 seconds (`:2869`). The nominal allowance is measured from adaptation start, not from the beginning of each thin step.
4. If none is available, the function returns `-1` with source `none`.

The optional/nominal paths capture `adaptStartedAt = CDbl(Now) * 86400#` at `RPX_Adapt.bas:1956` and subtract it from the same full-date expression at `:2973`. `Now` includes the date, so an ordinary midnight increases that expression rather than wrapping it to zero. The early-admission deadline test (`:2773-2837`, especially `:2803`) and `OptionalStageExpired` (`:183-192`) use this same full-date clock. They are midnight safe, although wall-clock adjustments can affect them.

The excerpt explicitly records `remaining_source=nominal` at rows 148 and 149. Thus neither a case limit nor an optional deadline supplied the logged HSD allowance there. The excerpt does not expose the original numeric case-limit setting, so the setting's exact value should not be inferred merely from its absence from this excerpt.

### HSD and DLL budgets

In rc9b, `RPX_DiagClock` delegates to `ClockSeconds` (`RPX_Diagnostics.bas:165-174`). Following `RPX_DiagStart` initialization at `:266-267`, this uses QueryPerformanceCounter/QueryPerformanceFrequency when available. Its fallback is full-date `CDbl(Date) * 86400# + Timer`, which also follows an ordinary midnight correctly, apart from the possible race between the separate Date and Timer reads and sensitivity to local clock changes.

`RPX_ThinSolveSeconds` (`RPX_Adapt.bas:2926-2945`) clamps a requested duration to remaining time minus the active side's completion reserve. The upper and lower reserves are independent of the thin-step diagnostic duration. The HSD start is captured with `RPX_DiagClock` (`RPX_Solve.bas:2101`). Its ordinary budget and thin hard-wall checks also use that clock (`:1968-2022`). The hard wall includes any endgame grace; it does not grant an extra reserve-sized solve.

The size/cone floor and the remaining-time fraction floor are allowances, not separate clocks (`RPX_Solve.bas:2108-2122`). The upper flat floor can then be reduced to current remaining time minus the upper reserve (`:2124-2147`). The final clamp is reapplied at `:2180`.

Native DLL admission and polling also consult `RPX_ThinSolveSeconds` (`RPX_AccuracyNative.bas:309-315`, `:349-364`). Its separate poll watchdog captures Timer at `:356` and corrects a negative difference by adding 86400 at `:368`. That watchdog already handled one midnight crossing; it was not an uncorrected endless-loop example.

### Round timing and extra-round protection

The measured lower-round cost that funds later extra rounds is captured with `lowerStarted = RPX_DiagClock()` at `RPX_Adapt.bas:3812` and computed at `:3820`. It does not consume the `thin_layer_step` text or its raw Timer subtraction. The cost model (`:3683-3701`) and extra-round admission (`:3705-3726`) use this measured value and the remaining-time function.

Upper protection for a pending extra lower round similarly uses `RPX_DiagClock` (`RPX_Adapt.bas:2982`, `:3018`). P06N row 32 explicitly records that this special protection was inactive for upper round 8: `active=0;reason=no_useful_lower_size`. Ordinary thin solver time protection remained active. The same row records positive measured `last_lower_seconds=8695.78883109998`, agreeing with the preceding lower-step report of 8695.7 seconds at row 31.

Round-count limits are counts; they are not Timer deadlines. Optional-stage vetoes, retry admission, extra-round affordability and solver walls use the clocks described above.

## Evidence from the supplied P06N excerpt

| TSV row | Timestamp / elapsed | Observation and implication |
|---|---|---|
| 31 | 23:31:07 / 15329.851955 | Lower round 4 reports positive `seconds=8695.7` and a valid result. |
| 32 | 23:31:07 / 15329.983270 | Upper protection begins with `active=0;reason=no_useful_lower_size`; measured lower cost remains positive. Upper round 8 work follows. |
| 148 | 23:43:49 / 16091.556228 | HSD has `remaining_seconds=2080.00000051968;remaining_source=nominal`. `time_fraction=0.15` produces `time_floor=312.000000077952`; the cone floor is 874.488 seconds. |
| 149 | 23:43:49 / 16091.556495 | The 1800-second upper size floor is reduced to `applied=1480.00000051968`, exactly the logged remaining allowance minus `reserve_seconds=600`. |
| 173 | 00:08:38 / 17580.600024 | HSD stops with `reason=thin_wall_limit;seconds=1489.04494940001;limit=1480.0019452197;iteration=146`. Its elapsed duration remains positive after midnight. The approximately nine-second overrun is consistent with checking around solver operations, not an extra 86400 seconds. |
| 179 | 00:08:38 / 17580.673555 | Solve 173 ends with positive `wall_seconds=1640.16597580002` and `HSD_SOLVE_BUDGET_EXCEEDED`. This whole-solve measure has a broader scope than the HSD elapsed measurement. |
| 180-181 | 00:08:39-00:08:47 | Further helper/auxiliary DLL calls have `packed_seconds=-1;state=0;iterations=0`: no actual native budget was packed. |
| 192, 211 | 00:09:33, 00:10:20 | Later original upper DLL attempts also have `packed_seconds=-1;state=0;iterations=0`. Their earlier `thin_layer_budget` events (rows 190, 209) show a proposed floor of 116.473551952785 seconds, not an allowance successfully granted after admission. |
| 195-201, 214-220 | 00:09:33, 00:10:20 | Subsequent HSD attempts are refused with `HSD_SOLVE_BUDGET_EXCEEDED`, zero iteration index and `thin_budget_no_iterate`; endpoints remain unchanged. |
| 222 | 00:10:21 / 17683.601660 | The independent raw Timer thin-step report is negative: `seconds=-84046.3`. |
| 223, 229 | 00:10:21-00:10:22 | The retained lower/upper values are 0.67486962890625 and 0.6914541015625. |
| 251 | 00:10:27 / 17690.141817 | `run_end` reports `PARTIAL_CERTIFIED` and `ADAPT_GAP_TARGET_NOT_MET`. |

There are **no explicit post-midnight `remaining_seconds` samples in this excerpt**. Rows 148 and 149 are before midnight. The positive wall expiry and subsequent refusals are the post-midnight evidence; a fabricated remaining-time sample would overstate the log.

For scale, subtracting the approximately 1489-second HSD interval from the prior 2080-second remaining allowance leaves about 591 seconds, below the 600-second upper reserve. This is an inference consistent with the subsequent refusals, not a logged remaining-time value. The prior allowance implies a nominal endpoint around 00:18:29 and a reserve-limited wall around 00:08:29, close to the actual 00:08:38 expiry. It shows no midnight-sized replenishment.

### Negative sentinel values are not negative elapsed times

`RPX_AccuracyNative.bas:279` initializes `packedSeconds = -1#`; only `:317` replaces it after a real allowance is packed. Earlier admission failures can leave this sentinel intact. `remaining_seconds=-1` likewise means unavailable/inactive where documented by the function contract. Neither value represents a negative measured duration.

The runtime verifier must reject negative elapsed values and actual negative budgets while identifying these narrowly defined inactive/not-started sentinels separately. Changing those sentinels would exceed this clock fix's scope. A positive proposed floor event also does not prove that a later DLL start was permitted.

The delivered runner follows the task's stricter acceptance rule: any negative numeric seconds/budget value, including a documented `-1` sentinel, fails log acceptance. Such a failure is reported for review and does not by itself establish a Timer defect. The production sentinel values are preserved.

## rc9c correction

`vba_rc9c/RPX_Clock.bas` is a new ASCII-only module whose VBA module name is `RPX_ClockModule`. It exposes one `Public Function RPX_Clock() As Double` returning seconds:

- On VBA7/Win64, the production path uses `GetTickCount64`, declared with `LongLong` only inside the Win64 conditional, divided by `1000#`.
- Other platforms use Timer plus module-level last-reading and day-offset state. Each observed backward wrap adds `86400#` before returning the reading.
- The fallback begins at a positive day offset, including when its first raw read is exactly midnight. This preserves the existing callers' use of zero as an uninitialized start sentinel.
- Elapsed/deadline start captures and later reads use this same helper. Diagnostic `ClockSeconds` delegates to it, so budgets, thin measurements and diagnostic timings share one epoch. Existing full-date elapsed calculations also use the helper; ordinary calendar timestamps and file-name timestamps remain calendar timestamps.
- Existing ad hoc one-wrap adjustments are removed from active elapsed paths because the helper already accounts for wraps. The original absolute Timer stage-log stamp remains a non-elapsed use.
- Mixed non-ASCII elapsed-report lines are retained byte-for-byte inside inactive `#If False Then` branches. Their active ASCII replacements preserve the displayed Unicode text using `ChrW$` and calculate elapsed seconds with the helper. An inventory/static scan must distinguish this preserved inactive reference text from executable elapsed Timer use.

Solver limits, reserves, floor formulas, retry counts, acceptance conditions, mesh policy and DLL binaries are not intentionally changed. The clocks' precision and origin do change: GetTickCount64 exposes millisecond units and platform tick resolution, coarser than the former QPC diagnostic clock. Very short phase measurements and threshold boundaries may therefore quantize differently. The old Now-based deadlines also had different resolution. Log formatting can still print more decimal places than the underlying clock resolves. Exact profiling-number equality is not an appropriate test; positivity, continuity, budget behavior and Fs agreement within `FS_TOL` are.

The new module keeps generic clock state and platform declarations separate from the existing native/DLL modules. To satisfy the explicit `LongLong` guard rule, two inherited local declarations in `RPX_AccuracyNative` are now enclosed by `#If Win64 Then`: the supported x64 declarations are unchanged, and the other branch uses 8-byte `Currency` storage for the handle/bit fields. This does not make the inherited whole project portable: other existing modules assume x64 Excel, `LongPtr`, PtrSafe declarations and x64 native APIs. The clock helper's fallback can be inspected or tested in isolation on other VBA targets.

### Default-off forced-wrap hook

`RPX_TestClockWrap` is permitted only while idle and when `TEST_INJECTION=1`. A positive argument below 86400, for example 120 seconds, forces the Timer fallback and rotates its seconds-of-day so that its next wrap occurs approximately that many elapsed seconds after configuration. This exercises the same backward-reading correction used by a real midnight. A zero argument disables the forced path. Module initialization leaves the hook off and inert.

Reconfiguration aligns the new mode to the previous clock reading so stored captures keep the same epoch. `RPX_TestClockStatus` provides evidence including active state, configured wrap time, observed wrap count, negative raw jump and nonnegative corrected step. The operator runner should configure each scenario while idle before beginning the run, then inspect this evidence after completion. A W1 run that never reaches the configured wrap does not demonstrate the fix.

### Fallback limits

The Timer fallback must be observed often enough to see every day boundary. It cannot infer whole unobserved days from two seconds-of-day readings; a gap of 24 hours or more is unsupported. Local clock adjustments can also look like backward wraps or forward time jumps, so the fallback does not provide the same wall-clock independence as GetTickCount64. The forced-wrap hook exercises wrap handling, not arbitrary clock adjustments, suspend/resume behavior, multiple unobserved days or every Office/platform combination.

## Verification status

This analysis is based on static source inspection and the supplied TSV evidence. No Excel/VBA execution, Python/reproduction-runner execution, native solve, build, git operation, network access or DLL change was performed for this analysis. The provided runtime plan and `repro_rc9c.py` are for later operator execution; their results are not claimed here. Static checks establish source-level properties, not an Office compilation or a successful numerical run.
