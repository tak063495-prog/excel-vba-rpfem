# rc9a static root-cause analysis

Line numbers below refer to the unchanged rc8 files in `vba/`, unless marked rc9a. No Excel, VBA, DLL, or network execution was performed. The supplied log has `line=0`, so a unique incident faulting instruction cannot be proved from these artifacts. This analysis distinguishes the established call path from ranked leaf candidates; it does not claim a reproduced runtime diagnosis.

## Established failure path

1. `RPX_Fs.RPX_Bound`, lines 181-182, resets the iteration count and starts a solve. On a newly assembled program, `RPX_Solve.RPX_ResetPreparedState`, lines 115-124, erases `xx`, `xy`, `XSlack`, and `XDual`.
2. `RPX_AccuracyNative.RPX_AccuracyDirectBound`, lines 929-936, temporarily creates a zero `xx` for auditing. Rejection restores the original absence of `xx`/slack at lines 1035 and 1071-1072. Native status 0 is not an accepted proof. An unsuccessful native call erases returned vectors in `AccuracySolveCore`, lines 251 and 350. A status-0 report therefore does not imply a usable VBA iterate.
3. `RPX_Solve.OptimizeHSDCore`, line 2103, calls `ThinHsdBudgetClamp`. The clamp can raise `5/HSD_SOLVE_BUDGET_EXCEEDED` at lines 1974 or 1983. This happens before HSD state allocation at line 2182 and before warm-cache setup at lines 2211-2226. Solve 176's zero iterations, no HSD initialization events, and immediate budget exception match this path.
4. `RPX_Fs.RPX_Bound`, lines 239-248, logs `solve_end NUMERICAL_FAILURE` and rethrows the original error. However, `FsValueCore.Failed`, lines 335-354, recognizes it as recoverable and enters `AccuracyFsRecover` at line 354. That recovery has no missing-iterate precondition. It attempts the original dual helper at line 992 and then primal recovery at line 1004. A secondary error here replaces the original recoverable budget error.
5. `TryFsValue.Failed`, lines 64-71, sees the replacement error 9, logs `stage=fs_trial`, and rethrows it because error 9 is not a numerical-recovery token. The same is true of `TryFsTrialPoint`, lines 120-128. This explains why the first visible error 9 is after `solve_end`, rather than inside Optimize/Bound.

Earlier failures, including solve 175, entered the HSD loop and had allocated state (56 iterations in solve 175). They did not exercise the same missing-vector recovery entry. Native status 0 alone is not established as the cause; the distinguishing combination is a rejected result followed by pre-allocation HSD failure and optional recovery without an iterate.

## Ranked exact statements and guards

| Rank | rc8 statement / procedure | Evidence and rc9a treatment |
| --- | --- | --- |
| 1 | `RPX_Fs.FsValueCore`, line 354: `If AccuracyFsRecover(...) Then Exit Function` | Confirmed recovery boundary after the original numerical error. rc9a lines 359-365 return the original recoverable error when this is a thin budget failure, iteration zero, and `xx` is absent or has the wrong shape. No optional recovery, sign, field, or bracket update follows this failure. |
| 2 | `RPX_AccuracyNative.RPX_AccuracyRecoverPrimal`, line 477: `If LBound(xx) = 0 And UBound(xx) = original.n - 1 Then center = xx`; line 479: `For i = 0 To UBound(center)` | `LBound`/`UBound` on an unallocated array raise 9. The original line 477 is inside `On Error Resume Next`, so that probe alone is **not proof of the escaping error**. The subsequent unguarded center-bound access is a secondary candidate if center storage was lost during the inline probe/copy path. rc9a lines 463-472 and 486-492 use a local, explicit allocation probe and copy only after sequential shape checks; the allocated zero center is preserved when no source exists. |
| 3 | `RPX_Audit.RPX_FieldGate`, line 12: `oldX = xx`, before the handler at line 16; restore line 31: `xx = oldX` | Unsafe snapshot/restore assumption when a separate helper produces a candidate but there was no previous iterate. This candidate requires helper success, so it is less consistent with a completely exhausted auxiliary budget. rc9a lines 6-16, 23-25 and 44-45 explicitly track whether a prior array exists and restore its erased state when absent. The candidate audit and its acceptance criteria are unchanged. |
| 4 | `RPX_AccuracyNative.RPX_AccuracyDirectBound`, lines 930-933: inline bounds probes and snapshot copies | Locally suppressed probes are not the logged escaping error, because they occur before `direct_dll_end`. They are a related missing-state hazard. rc9a lines 943-950 use the same explicit shape helper and retain the original shape predicates. |

The first-ranked fix closes the incident's entire optional-recovery path before any of the candidate leaf statements. No global error-9 whitelist was added. Runtime instrumentation or a numbered VBA reproduction would be needed to identify one unique original leaf instruction.

## Paths checked and excluded

- Thin HSD warm capture/apply/key functions have local invalid-state handlers. Warm application is not reached when the initial HSD budget clamp raises. Their bounds checks do not explain the escaping post-solve error here.
- `FsArrayReady` and `ThinFsCompleteField` already handle absent arrays locally. `ThinFsObserve` is reached only after `FsValueCore` returns, and requires a complete physical field before copying one.
- Normal root-loop assignments at `RPX_Fs` lines 931/933 occur only after `TryFsValue=True`. Failed trials leave `a`, `b`, `fa`, `fb`, endpoint fields, and saved brackets unchanged. The final endpoint copies in `RPX_FsPair` lines 959/965 and `Fix8ChainStep` line 3417 are downstream of a successful endpoint return, not the observed escaping `fs_trial` error.
- Original dual/primal helper audits are already conditional on a finite helper result. Their tolerances and proof acceptance were not changed.

## Why confirmed bounds disappeared, and preservation

`RPX_Adapt.Fix8ChainStep`, lines 3475-3505, classifies an unexpected evaluation error as fatal, restores its baseline, and rethrows it at lines 3470-3472. The caller has no encompassing handler after `FinalCandidateDone` (lines 2137-2167). `RPX_Control.RPX_Run.FailedCleanup`, lines 300-308, then clears result cells B4:D30. Retaining witness objects inside the thin round was insufficient to produce a worksheet result.

rc9a adds an outer adaptive wrapper and `RPX_AdaptPreserveConfirmed` (Adapt lines 1875-1932). Recovery uses the current audited best states and matching current-run `ADAPT_BEST_STATE` certificates; it does not gate on `currentPairValid`, which a failed candidate can clear while the best witnesses remain valid. It checks run, inputs, policy, owner metadata, finite ordered positive Fs, and existing input assertions. It clears thin budgets/cache/search activity, restores the best upper witness's own model and final bound fields, marks the result incomplete, and logs `adapt_fatal_preserved` with the original error and witness values/elements. It never changes either best state or accepts a new endpoint.

The wrapper returns normally to the original result-writing path. A one-attempt Control recovery branch also covers errors after adaptive return, including the final output hook. Existing Japanese/source lines in Control are unchanged; only ASCII additions were necessary to catch this outer scope. Successfully recovered Control error flags are cleared after the original error has been retained in the diagnostic and search note.

Cancellation, changed inputs, stale/mismatched certificates, unavailable/corrupt witness storage, or a repeated restoration/output failure retain the original failure handling. A persistent worksheet write failure cannot be made writable by retrying; recovery is deliberately bounded to one attempt. See `TEST_PLAN.md` for reproducible, default-off hooks and `static_checks_rc9a.txt` for validation limits.
