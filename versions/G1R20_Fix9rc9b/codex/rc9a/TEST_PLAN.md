# rc9a test plan

Only static checks were run for this delivery. No workbook, Excel, VBA procedure, or native DLL was executed. The following are operator-run tests for a disposable workbook copy after importing the changed modules.

## Setup

Use a small existing frictional P02A_THIN-like or D07A case with roughly 256-400 target elements, no rigid bodies, no coupled water, and constant body force per material. Use its established working solver settings. Start with:

| Setting | Value |
| --- | --- |
| MODE | Fs |
| ADAPT | 1 |
| TEST_INJECTION | 1 |
| ACCURACY_DLL_RECOVERY | 1 |
| ADAPT_PRECISION_FROM_START | 1 |
| ADAPT_THIN_LAYER | 2 |
| ADAPT_THIN_UPPER_ROUNDS | 1 |
| ADAPT_THIN_UPPER_CAP | 800 |
| ADAPT_THIN_UPPER_ROUND_MAX_ELEMENTS | 800 |
| ADAPT_THIN_EARLY_UPPER_CAP | 800 |
| ADAPT_THIN_LOWER_ROUNDS | 0 |
| ADAPT_THIN_LOWER_EXTRA_ROUNDS | 0 |
| ADAPT_THIN_MID_SEPARATE | 0 |
| GAP | 0.000001 |

Keep the existing FS_TOL, audits, and numerical tolerances. The small GAP requests continuation; it does not weaken a proof gate. Thin mode 2 forces detection but still requires valid geometry and material/body-force scope. Confirm `thin_layer_begin` appears and an upper candidate reaches endpoint evaluation. If the coarse pair already meets the requested gap, or the geometry cannot build a candidate, the thin hook will not fire. `final_accept` below remains available for the broad preservation test. Runtime depends on the case and machine; these hooks do not require P06N or a long wall-budget wait.

The existing mechanism accepts TEST_INJECTION=0/1 only. Reuse value 1 and explicitly arm one hook before running, while RPX_Busy=False. Do not add new numeric injection values. In the VBA Immediate window, run `RPX_TestDisarm` before each scenario. Each armed hook is consumed once; `skip=0` targets the first matching call and `skip=N-1` the Nth call.

## A. Missing iterate / exhausted HSD budget

Arm, then start the normal analysis:

```vb
RPX_TestDisarm
RPX_TestArm "thin_upper_no_iterate", 5, "HSD_SOLVE_BUDGET_EXCEEDED", 0
RPX_Run
```

This new hook runs after assembly and before solver dispatch, only for a thin upper trial. It bypasses the solvers, erases the four iterate arrays, sets iteration zero, and raises the existing budget token. It simulates the rejected-DLL/no-result plus pre-allocation HSD state; it does not execute a fake native solve or fabricate `direct_dll_end`.

Expected log:

- `test_injection stage=thin_upper_no_iterate;number=5`, then `test_no_iterate` with `simulation=solver_bypassed;simulated_dll_status=0;dll_result_absent=1;hsd_allocated=0`.
- `error_origin stage=bound` and `solve_end status=NUMERICAL_FAILURE;last_iteration_index=0` for the injected solve.
- For q8 accuracy-enabled root search, `fs_numerical_unknown reason=thin_budget_no_iterate;endpoints_unchanged=True`, followed by the original recoverable error at `stage=fs_trial`, never a replacement error 9.
- The failed trial changes no bracket endpoint or best witness. Existing retries may subsequently finish the round successfully; otherwise the existing failed/unknown round path retains the old witnesses. Do not expect every one-shot injection to force the whole round to fail.
- The normal output path writes finite ordered B4:C4. An unmet gap/search tolerance uses the existing partial status and `run_end PARTIAL_CERTIFIED`.

An all-zero-friction endpoint uses allowRecovery=False. It still handles the injected budget error through existing round recovery, but may not emit the new `thin_budget_no_iterate` diagnostic. Use a frictional Fs root search to verify that exact guard.

## B. Unexpected error 9 inside a thin round

```vb
RPX_TestDisarm
RPX_TestArm "fix8_thin_upper_1", 9, "RC9A_THIN_FATAL"
RPX_Run
```

This uses the existing candidate hook. With separate mid enabled, `_0` is the mid step and `_1` remains the first regular upper round. Expected: `thin_layer_rejected ... reason=fatal`, followed by `adapt_fatal_preserved number=9 ... message=RC9A_THIN_FATAL`. There must be no subsequent optional chain/accuracy/fan solve. Record the last audited lower/upper values and own element counts before injection; they must equal those in the preservation event and the result worksheet. Fs remains in B4:C4; search incomplete is true and the run ends PARTIAL_CERTIFIED rather than NUMERICAL_FAILURE. No trial at the injected candidate is accepted as a new bound.

## C. Late adaptive and outer output errors

Run separately:

```vb
RPX_TestDisarm
RPX_TestArm "final_accept", 9, "RC9A_FINAL_FATAL"
RPX_Run
```

Then separately:

```vb
RPX_TestDisarm
RPX_TestArm "output", 9, "RC9A_OUTPUT_FATAL"
RPX_Run
```

Both must emit `adapt_fatal_preserved`, retain the same current-run best pair, and complete the normal partial result commit. The output-hook case verifies the Control-level Resume path after RPX_RunAdapt has returned. Its successful recovery clears RPX_LastError and the failure number/source; the error remains in the diagnostic/search note. A second genuine restoration/output failure must take normal failed cleanup rather than loop.

## D. Negative and regression checks

1. Disarm and set TEST_INJECTION=0. Repeat the small case with identical settings in rc8 and rc9a. Compare accepted bounds, witness mesh sources/counts, Fs calls, audits, and normal branch/solver sequence. No new injection or preservation diagnostic should occur. Ignore wall-time/process-memory fields.
2. Arm error 9 at `fs_value` before any pair is confirmed. It must remain a fatal failure with no fabricated Fs pair or preservation event. The new recovery must not use a previous run's arrays/certificates.
3. Arm cancellation 18 at `fix8_thin_upper_1`. Cancellation must retain existing cancelled-run behavior and must not emit `adapt_fatal_preserved`.
4. Use the existing output mutation hook (`RPX_TestArm "output", 5, "MUTATE_MATERIAL"`) in a disposable copy. Existing input assertions must prevent a result commit; no stale pair may be published. Restore that copy's inputs afterwards.
5. Repeat a normal no-error nonadaptive run and a non-thin adaptive run. Their bounds and accuracy decisions must match rc8.
6. Finally reproduce the supplied P06N settings if long-run validation is desired. Zero-iterate budget failures must remain numerical unknowns, and any late unexpected adaptive error must retain the last matching audited best pair in normal partial output.

After testing, call `RPX_TestDisarm` and restore TEST_INJECTION=0 and all production settings.
