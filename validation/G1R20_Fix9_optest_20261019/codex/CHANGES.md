# Fix9 candidate changes

Only `vba/RPX_Adapt.bas`, `vba/RPX_Solve.bas`, and `vba/RPX_AccuracyNative.bas` are modified. No modules were added. The candidate changes thin-layer mesh construction, recovery from a failed lower candidate, and solve time budgets. Existing q8 audits, `RPX_FsEndpoint`, `RPX_ActualRootBracket`, `SaveState`, `RPX_CertFromAdapt`, `RecoverEquilibrium`, certificate rules, and `FS_TOL` are unchanged.

## Scope and cleanup guards

`Fix8ThinLayerChains` retains these entry guards before reading Fix9 settings or calling any new helper:

```vb
If mode <> "Fs" Or Not currentPairValid Or Not RBestLower.complete Or Not RBestUpper.complete Then Exit Function
If Not Fix8ThinMode(mode, mask) Then Exit Function
```

The geometry helpers and `Fix9LowerStep` are private and called only through this driver/`Fix8ChainStep`. `Fix8ChainStep` additionally requires a valid pair, complete q8 lineage, its valid own root, no incomplete search, `RR=0`, and no water coupling. It sets `RPX_ThinChainBudgetActive=True` and `RPX_ThinChainUpperActive=(side="upper")` only after installing its error handler. No shared flag or budget value is set by the driver or by a non-thin path.

`Fix9ClearBudgets` resets both flags and every new shared budget parameter. `CandidateDone`, `CandidateFailed`, and `RestoreFailed` call it; an auxiliary restoration failure also calls it before propagating. Applicability exits occur before flags are set. Cancellation and fatal errors use the existing baseline restoration/error path and propagate, rather than becoming retries. Outside these chain steps the HSD hook is guarded off, and the DLL hook is guarded off; ordinary Fix8 solver decisions and limits are retained.

## Six requested changes

1. **Separate first mid-line step.** `RPX_Adapt.bas`: `Fix8ThinLayerChains`, `Fix8ChainStep`, and new `Fix9BuildMid`. With `opt.midSeparate`, each enabled lineage first takes a round-0 candidate containing only `RPX_TransverseEdges`/`RPX_SplitEdgeSet`; it has no indicator edges, fan, or flips. It uses the existing own-mesh root/acceptance path, and result 1 or 2 advances the lineage. The driver then passes `withMid=False` to later steps. With `ADAPT_THIN_MID_SEPARATE=0`, the original first-step merged mid-line construction remains in the upper/lower builders. The new branch is reachable only after the driver guards above.

2. **Upper indicator on the lower lineage mesh.** `RPX_Adapt.bas`: `Fix8BuildThinLower`, new `Fix9AuxUpper`. Guard: `If opt.lowerOwnUpper Then`, inside the chain. One upper solve at `opt.lowerOwnFactor*lineage.value` produces velocity for `RPX_ThinUpperScore`, with no point location. Failure returns to the existing best-upper density/location calculation through `If Not ownUpper Then` and emits `thin_layer_aux`.

3. **Fresh upper velocity before flips.** `RPX_Adapt.bas`: `Fix8BuildThinUpper`, `Fix9AuxUpper`. Guards: `If passes > 0 Then` and `If opt.flipFresh Then`, inside the chain. After splitting and before `RPX_FlowMeshCandidate`, one upper solve at `lineage.value` supplies velocity. `If Not fresh Then` retains `RPX_ProlongUpperVelocity` as fallback. Every flipped candidate still requires its own ordinary root and acceptance.

4. **Bounded lower retries.** `RPX_Adapt.bas`: `Fix8ChainStep`, `Fix8ThinLayerChains`, new `Fix9LowerStep`. A step now returns `stopReason` alongside 0/1/2. Only lower result 0 with `root_failed` or `evaluate_failed` is retried, from the unadvanced lineage, multiplying theta by `opt.retryThetaFactor` each time. The default permits one retry; settings allow 0 through 3. `cap`, `no_edges`, `geometry`, `crossed`, cancellation, fatal errors, and applicability failures do not trigger retries. The wrapper is private to the guarded thin driver; upper stop behavior is unchanged.

5. **Per-solve HSD budget floor.** `RPX_Solve.bas`: new shared chain flag/parameters and the narrowly guarded block in `OptimizeHSDCore`. Exact hook guard: `If RPX_ThinChainBudgetActive Then`. After compilation, the current model's `XB` cone count determines `min(cap, factor*(base+per_cone*XB))`; this raises `solveBudget` only when greater. The ordinary `RPX_HsdBudgetFloor` mechanism remains intact and is never assigned by Fix9. The extension is local to that solve and logs ordinary/floor/applied seconds.

6. **Upper DLL time scaling.** `RPX_AccuracyNative.bas`: new parameters/helper `RPX_ThinUpperDllSeconds` and one call in `RPX_AccuracyDirectBound`. Exact call guard: `RPX_ThinChainBudgetActive And RPX_ThinChainUpperActive And mode="upper"`. The helper also exits unless both flags are set and cones are positive. It returns `max(ordinary, min(cap, factor*ref_seconds*(cones/ref_cones)^exponent))`, using `original.coneCount-1` to exclude the exported equality block. Lower root solves keep the ordinary DLL limit. Auxiliary upper solves temporarily set the upper flag and restore it, so both fresh-flip and lower-indicator upper solves receive the same scaling while remaining inside the chain.

## Auxiliary isolation and implementation choices

`Fix9AuxUpper` first requires `RPX_ThinChainBudgetActive`. It assembles/exports the current mesh's upper problem and calls `RPX_AccuracySolve` exactly once. It never calls `RPX_Bound`, a root solve, recovery, `SaveState`, or certificate acceptance. Only a finite velocity of the expected size is returned. The helper saves/restores the audit snapshot, q/target/tolerance, root-search metadata/calls, native diagnostics, strength/program/order state, objective/status/physical-gate/work values, raw lower audit values, lower/upper Fs and fields, current iterate, and prior upper-budget flag. Compiled work is invalidated as in the chain's restore path. It does not assign `RBestLower`, `RBestUpper`, `RCurrent`, certificates, or bracket structures. Each ordinary chain candidate retains the existing save/restore block and witness gates.

Choices and limitations:

- Mid-line candidates are round 0 and do not consume configured upper/lower rounds or stall counts. A `no_edges` mid candidate is skipped without a root attempt, but its call is included in diagnostic `steps`. The regular candidate may run in the same outer round. Disabled chains receive no extra mid step.
- Every lower attempt, including retries and mid-step retries, counts as a step. Retries are additional calls within the same configured round. Retrying a mid-only failure repeats the same transverse geometry; reduced theta affects only indicator candidates. This keeps the required mid-only operation pure while allowing transient solve recovery.
- Auxiliary solves are DLL-only: no HSD fallback or internal retry is added, preserving the single-solve requirement and limiting indicator cost. Missing DLL, rejected export, solve failure, timeout, invalid velocity, or exception falls back to Fix8 geometry and is logged; cancellation propagates after restoration.
- Auxiliary acceptance conservatively requires native solver status 1 (`Solved`), as well as successful return, elapsed time within limit, and finite velocity. The native implementation/status mapping is not supplied here and this mapping has not been runtime verified. Other statuses fall back safely. Auxiliary fields are indicators only and never become witnesses.
- New HSD/DLL caps cap the added floor, not an existing larger ordinary limit: neither change lowers a user's ordinary budget.

## New settings

All values are read only inside the guarded driver through `Fix9ReadOptions` and `Fix8Number`, which rejects nonfinite/out-of-range values. `Fix9Integer` additionally rejects fractional values before `CLng`; booleans, retry count, and reference cone count use it. Bounds below are inclusive. All 15 resolved values are appended to `thin_layer_begin`.

| Setting | Fix9 default | Valid range/type |
|---|---:|---|
| `ADAPT_THIN_MID_SEPARATE` | 1 | integer 0..1 |
| `ADAPT_THIN_LOWER_OWN_UPPER` | 1 | integer 0..1 |
| `ADAPT_THIN_LOWER_OWN_UPPER_FACTOR` | 1.08 | 1..2 |
| `ADAPT_THIN_FLIP_FRESH` | 1 | integer 0..1 |
| `ADAPT_THIN_LOWER_RETRIES` | 1 | integer 0..3 |
| `ADAPT_THIN_RETRY_THETA_FACTOR` | 0.5 | 0.000001..1 |
| `ADAPT_THIN_HSD_BASE` | 20 | 0..3600 |
| `ADAPT_THIN_HSD_PER_CONE` | 0.0037 | 0..1 |
| `ADAPT_THIN_HSD_FACTOR` | 1.5 | 0..100 |
| `ADAPT_THIN_HSD_CAP` | 3600 | 1..3600 |
| `ADAPT_THIN_DLL_FACTOR` | 3 | 0..100 |
| `ADAPT_THIN_DLL_REF_SECONDS` | 3.45 | 0.000001..600 |
| `ADAPT_THIN_DLL_REF_CONES` | 22741 | integer 1..2147483000 |
| `ADAPT_THIN_DLL_EXPONENT` | 1.2 | 0..4 |
| `ADAPT_THIN_DLL_CAP` | 600 | 0.000001..600 |

Existing thin-chain settings and defaults, including round counts, theta, caps, start round, and flip passes, remain unchanged.

## Diagnostics

| Event | Added/extended information |
|---|---|
| `thin_layer_begin` | All new settings; `mid_counts_as_round=0`; `retries_count_as_steps=1`. |
| `thin_layer_step` | Explicit stop reason, evaluated/result; evaluated steps include mid-only, theta, Fs, root bracket/calls; mid geometry records zero indicator/fan/flip counts. |
| `thin_layer_rejected` | Error stage and reason, retained witnesses, error message; cancellation/fatal classification precedes propagation. |
| `thin_layer_aux` | Purpose, Fs, elements, success, reason, time limit, `witness=0`; emitted for success and fallback. |
| `thin_layer_retry` | Lower round/retry, prior failure reason, reduced theta, unchanged lineage Fs/elements, mid-only flag. |
| `thin_layer_budget` | Solver/mode, current model cones, ordinary seconds, computed floor, applied seconds. |
| `thin_layer_end` | Existing pair/mesh/GAP reporting; `steps` includes extra mid candidates and retry attempts. |

## Validation

Review is static only. Excel, VBA execution, and long computations were not run, as required. No native runtime or VBA compiler acceptance is claimed. See `diff_summary.txt` for changed modules and line ranges and `ANALYSIS.md` for the supplied result evidence and its limitations.

Static checks passed: exactly 3 of 63 modules changed; the other 60 are byte-identical to `fix8_orig_vba/`. Every existing procedure outside the four permitted thin-chain functions and the two solver hook functions is byte-identical to its original. All three modified modules are pure ASCII, retain `Option Explicit`, and preserve the original newline conventions (including the original 23 LF-only endings in `RPX_Solve.bas`). Maximum physical line lengths are 373 (`RPX_Adapt`), 424 (`RPX_Solve`), and 569 (`RPX_AccuracyNative`), below VBA's 1023-character limit. Source review checked declarations, flag/error cleanup, retry classification, and auxiliary state isolation. The line-diff reconstruction was checked against every matching line and both source lengths before writing `diff_summary.txt`.
