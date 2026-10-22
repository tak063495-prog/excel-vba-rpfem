# G1R20 Fix9 rc9b design (revised task)

rc9b is rc9a plus literal original-upper DLL equilibration and the dual
logging correction. The fresh `vba_rc9b/` copy is derived from read-only
`vba/`. Only RPX_AccuracyNative.bas, RPX_UI.bas and RPX_P4.bas differ.
The numerical change is confined to native settings selection;
existing acceptance, physical audits, root logic, fallback schedules, budgets,
DLL declarations and DLL binaries are unchanged. No Excel/VBA run, workbook
import, build, native replay, git operation, network access or DLL change was
performed. Static checks are source checks, not a VBA compilation result.

## The sign finding and removal of early rejection

An audited objective below the actual upper target `0.999999` is the
certified upper-witness side of this search:

- `RPX_AuditUpperCore` computes `RValue = RDissipation - RFixedWork` and checks
  agreement with `XObjective` within `0.000002`
  (`vba/RPX_Audit.bas:307-311`).
- `FsValueCore` returns `RPX_Bound(...) - target`
  (`vba/RPX_Fs.bas:286`), with upper target `0.999999` (`:611`).
- A negative residual updates endpoint `b` (`:969-973`). The final upper
  endpoint restores `b` and its field and reruns `RPX_AuditUpper`
  (`:980-983`). A nonnegative residual updates `a`.

The revised task confirms this interpretation from historical logs: cap3's
final upper Fs `0.692342` came from a helper-status-5 trial that HSD certified
with objective `0.99935531472274` (solve 186 in `st5scan.txt`). An early reject
based on a below-target objective would discard a real upper witness or give
it the opposite search sign. That feature is removed completely: there is no
new rejection setting, margin, status, predicate, event, root branch,
interpolation rule, endpoint cleanup or acceptance path. `RPX_Fs.bas` remains
byte-identical to rc9a. Existing rc9a search-only statuses retain their original
scope; helper status 5 supplies no new sign or certificate.

## Upper equilibration setting

`ACCURACY_DLL_UPPER_EQUILIBRATE` accepts only finite 0 or 1. An absent row
uses 1, the settings row defaults to 1, and both the common accuracy preset
and improvement preset explicitly set 1. The Speed2 preset inherits the
common reset, so E1 does not retain a previously saved setting of 0. Setting 1 overrides the existing packed equilibration bit only for a
literal current original upper primal LP, including its existing eligible
status-5 retry. The discriminator requires upper mode, `CurrentProgram`, zero
quadratic diagonal and cost equal to the original cost. It therefore excludes
auxiliary cap/QP programs even though some inherit upper mode.

Setting 0 disables this override and reproduces rc9a settings selection. It
does **not** force the ABI bit to zero: rc9a thin original solves use
`ADAPT_THIN_DLL_EQUILIBRATE` (default 1), ordinary primal solves use
`ACCURACY_DLL_EQUILIBRATE` (default 0), and eligible legacy retry selection is
retained. This distinction matters when comparing E1 with E0 on P02A_THIN:
both may legitimately pack 1 for upper thin original LPs.

Lower primal, auxiliary and original-dual-helper calls retain their existing
selection. The helper is explicitly labelled `original_dual_helper` and still
uses `ACCURACY_DLL_DUAL_EQUILIBRATE` (default 1). The new setting does not
modify solver duration, iteration limit, tolerances, regularization, cone
scaling, HSD, equality repair or objective/physical acceptance thresholds.

The supplied replay evidence covers 76 upper dumps, with native status 4 in
both variants. Total wall time was 408 seconds without equilibration versus
264 seconds with equilibration: 144 seconds less (35.3%). The supplied
base-minus-e1 lambda difference has median +0.00021 and maximum +0.058,
showing an upward bias in those unequilibrated samples. Median original dual
stationarity improves from 0.000022 to 0.00000072. These are supplied replay
results, not tests run here, and are not mathematical accuracy guarantees or
substitutes for the existing audits.

## Every existing original-primal acceptance gate

References below use the unchanged rc9a source line numbers in `vba/`. The
ordinary direct branch, before any original-dual audit or helper fallback,
sets `kind="original_primal"` at `RPX_AccuracyNative.bas:1007` only after all
applicable gates below. rc9b does not add a parallel acceptance route.

1. **Direct scope and current compiled problem.** `RPX_Bound` must allow the
   direct path and compile the current original problem (`RPX_Fs.bas:203-207`).
   `RPX_DirectEnabled` requires DIRECT_DLL=1, ACCURACY_DLL_RECOVERY=1, active
   Fs, q=8, non-affine mode and no rigids (`RPX_AccuracyNative.bas:898-903`).
   The ordinary LP branch must be reached: an already accepted helper-first
   guide returns through its existing route before the LP. The existing
   direct seconds value must be finite, positive and <=600
   (`:931-934`), and existing time/cancellation guards still apply.
2. **Literal export and identity.** `RPX_AccuracyExport` must succeed
   (`:138-208`): no cancellation/affine mode, active Fs, raw audit enabled,
   q=8, RR=0, upper or lower mode matching the order mode, reference_total
   phase, finite positive Fs equal to the active strength factor, ready
   non-affine compiled original problem, matching dimensions, input epoch,
   model, mode, load phase, subdivisions and strength factor. Dimensions must
   be positive; original-cost shape and entries, original equality rows,
   cone dimensions (1 or 3), offsets, indices and coefficients must be valid
   and finite; sparse limits, row accounting and `CurrentProgram` must pass.
   The LP export has P=0 and literal original cost.
3. **Current context remains valid.** `CurrentProgram` (`:93-109`) requires
   active Fs, raw audit, non-affine/no-rigid state; the same exact model,
   input epoch, q=8, reference_total phase, active mode, policy and Fs;
   current dimensions/equality count; ready original compilation with the
   same epoch, model, subdivisions, mode and finite matching strength
   factor; finite positive XQScale; and finite unchanged normalized/original
   cost entries. `RPX_AssertInputs` at the physical gate also rejects input
   changes and cancellation (`RPX_Audit.bas:22`, `RPX_Guard.bas:451-459`).
4. **A finite native candidate is actually returned.** The DLL must be
   available at the unchanged workbook path with ABI 1, settings/create/start/
   poll/result operations must succeed, worker state must be completed (2),
   the call must not be cancelled, both returned vectors must contain only
   finite values, and cleanup must succeed (`:245-360`). Missing or failed
   vectors go to existing recovery/fallback. A failed eligible retry restores
   its first finite candidate rather than discarding it (`:373-406`). Native
   status alone is never a feasibility certificate.
5. **Original primal algebraic audit.** `RPX_AccuracyPrimalAudit`
   (`:407-438`) requires `CurrentProgram`, candidate bounds exactly 0..n-1,
   finite entries and finite compensated objective/row calculations using
   `originalCost`. Maximum original equality residual must be <=0.0000001
   and maximum original linear/SOC cone violation must be <=0.0000001.
   Overflow, invalid array shape and nonfinite accumulation fail the audit.
6. **Existing thin upper status restriction.** With both thin-chain budget
   and upper-chain activity, only native status 1 or 4 can continue as an
   upper primal candidate (`:993-995`). A finite status-5 original-primal
   vector does not pass this gate. The helper's status does not satisfy or
   replace this original-primal status check.
7. **Full physical field audit at the same Fs and q=8.** The candidate must
   pass `RPX_FieldGate(mode,candidate,factor,8,objective,...)` (`:996`). It
   installs this candidate with the expected original objective and invokes
   `RPX_AuditUpper(True)` (`RPX_Audit.bas:17-54`). That audit recomputes
   dissipation, reference/fixed work and all physical constraints from the
   field (`:204-312`): rigid fixation/motion consistency where applicable,
   element volumetric/flow constraints, bonded continuity, unbonded normal/
   tangential jump constraints on every q=8 edge subinterval, and fixed,
   roller and rigid boundary conditions. It checks reference-work
   normalization to 1, includes the existing precise zero-cost work
   calculation when applicable, and sets `RValue=RDissipation-RFixedWork`.
   The maximum physical error RAudit must be <=0.0000001 and
   `Abs(RValue-XObjective)` must be <=0.000002 (`:303-311`). Any physical
   rejection or raised error prevents acceptance.
8. **Measured upper value threshold.** The physically measured value, not
   just DLL lambda, must be <=0.999999 (`RPX_AccuracyNative.bas:997-1008`).
   The lower homogeneous normalization branch is lower-only and cannot
   authorize an upper result.
9. **Commit and repeated physical audit.** Restore must complete without
   cancellation/error (`:1050-1055`). For accepted original_primal, the code
   installs the candidate and original objective, runs `RPX_AuditUpper`
   again, then labels `ORIGINAL_ACCURACY_UPPER_WITNESS` and performs the
   existing optional proof capture (`:1057-1064`). A raised commit-audit/
   capture error rejects the direct attempt through its existing handler.
   Successful return lets `RPX_Bound` exit before `RPX_OptimizeBound`/HSD
   (`RPX_Fs.bas:207-212`), and `FsValueCore` recognizes the physical witness
   (`:307`). A final returned upper endpoint still passes the existing
   root/bracket and fresh physical endpoint audit (`:977-983`); adaptation
   certification is unchanged.

Optional existing upper recenter recovery can also set original_primal, but
it is separate from the proposed seven ordinary-LP takeovers. It still runs
original primal audit, thin status 1/4 restriction, the full q=8 field gate,
finite measured/audit checks, measured<=0.999999 and recenter context assertion
(`RPX_AccuracyNative.bas:1031-1048`). Its cap/QP settings are unchanged.

## Can the seven replayed helper-status-5 trials bypass HSD?

Yes, **conditionally through the existing original_primal branch**. The
supplied cap3 replay trials 118, 127, 130, 139, 157, 166 and 181 have native
original-primal status 4 in both runs. Their base lambda is >=0.999999 while
equilibrated lambda is approximately 0.998803..0.999905, below the target.
Status 4 meets the existing thin upper status gate. Their equilibrated
objective clears the numerical threshold by at least about 0.000094; if the
unchanged physical objective-agreement gate passes, even an upward
0.000002 difference leaves the measured value below target. The threshold
itself should therefore no longer block these candidates.

The missing proof is the rest of the original/physical audit, especially the
returned vector's original equality/cone residuals and complete same-Fs q=8
field audit. Replay lambda and native status do not establish these. If all
listed gates pass, the ordinary branch sets original_primal before it calls
the ordinary original-dual/helper fallback, and `RPX_Bound` skips HSD. If an
existing helper-first schedule already ran, its rejected helper cost remains,
but a passing subsequent original LP can still avoid HSD. If any gate fails,
existing recovery and HSD fallback remain in effect. No helper failure is
reinterpreted as a primal witness or search sign.

## Defensible expected effect

Summing the listed integer row durations in `st5scan.txt` gives:

| Historical helper-status-5 HSD fallback group | Cases | HSD seconds |
| --- | ---: | ---: |
| cap3 | 36 | 13115 |
| meas | 4 | 6017 |
| P06A (rc8/a log) | 3 | 489 |
| Total | 43 | 19621 |

There are 40 OPTIMAL rows and 3 NUMERICAL_FAILURE rows. The printed per-group
subtotal lines differ slightly from their integer row sums; their sum is
19624. The revised task's 19621 seconds agrees with the listed rows, which
are used consistently here.

The seven specifically replayed candidates account for historical HSD times
75, 90, 91, 122, 223, 363 and 1198 seconds: **2162 seconds (36.0 minutes)**.
They are 7/43 cases (16.3%) and 11.0% of historical helper-status-5 HSD time.
They form the evidenced plausible takeover set, conditional on the unchanged
full gates. Zero of those seven is proven to pass all q=8 gates by the data
supplied here; consequently no nonzero avoided-HSD count is guaranteed.

The remaining **36 cases / 17459 seconds** have no supplied matching e1
primal/physical audit evidence. Their eligibility is unknown, including the
three HSD numerical failures. Historical below-target HSD objectives show
where HSD found real upper witnesses; they cannot prove that a DLL vector
will pass. Taking over all 43 would avoid at most the historical 19621 HSD
seconds (5.45 hours), an opportunity ceiling, not an expected saving. The
76-dump DLL wall improvement is a separate supplied result; do not add it to
historical avoided-HSD estimates or assume unchanged future search paths.

## Dual diagnostics and phase attribution

The logging correction follows `dual_prev/PROPOSED_FIX_dual.diff`, retaining
all numerical dual gates. Per-trial and helper entry clear dual residuals,
objective, validity and call attribution; an unavailable or incomplete audit
has `dual_audited=False`. Final log output renders unavailable dual metrics,
status and role as zero/empty through diagnostic locals. This validity rule
does not clear partial operational globals on incomplete audits, preserving
the existing repair scheduling; entry resets remain explicit. A completed but
rejected audit can legitimately have `dual_audited=True`. After a failed helper-first attempt, reset again
before the ordinary LP so a primal-only result cannot inherit helper fields.
Native status and iterations reset before each new native call. Each call
has a monotonically assigned ID; failed-retry restoration retains the origin
ID of the candidate being audited.

`accuracy_dll_call` logs `call_id`, `program_mode`, `program_role`,
`original_primal`, `upper_equilibrate_setting`, `thin_original`,
`thin_upper_retry`, dimensions, requested/packed seconds, iteration limit,
actual packed `equilibrate`, worker state, poll/result status, iterations,
cancellation and finite candidate availability. A completed-call event is
emitted on the existing cleanup path. Early exits before that path have no
completed event and must not reuse another call's status.

`accuracy_dual_vector` identifies native origin and candidate role
(original_dual, helper_primal, repair_source or repair_corrected), original
cost infinity norm and coordinate, dual infinity norm, stationarity maximum
coordinate, signed stationarity at the cost coordinate, original dual
objective and cone residual. Its derived `aty_inf_derived` subtracts cost
from the compensated residual; it is not an independently accumulated A' y.
Its validity flag covers diagnostic subtraction overflow without changing
certification. The exact existing dual test remains stationarity<=0.0000001,
cone<=0.00000001 and original objective>=threshold; upper threshold remains
1.000003, lower remains -0.999998.

`direct_dll_end` snapshots the original primal phase under `primal_call_id`,
`primal_status`, `primal_equilibrate`, `primal_audited`, `primal_audit_pass`,
`primal_raw_eq`, `primal_raw_cone`, `primal_objective` and `primal_role`.
Legacy `solver_status`, `raw_eq` and `raw_cone` consistently alias that saved
primal phase. Separately named `dual_audited`, `dual_call_id`, `dual_status`,
`dual_role`, `dual_stat`, `dual_cone` and `dual_objective` identify the dual
audit. `last_call_id` records the latest executed call, while
`candidate_call_id` records the retained candidate origin; these may differ
after a failed retry. A helper status cannot masquerade as a primal status.

With existing opt-in ACCURACY_DLL_CAPTURE=1, raw helper/original dual vectors
are captured before repair via existing AC18P001 diagnostic writers. Their
candidate_only=1 and physical_bound=0 markers remain explicit; they are not
newly accepted physical proofs. Capture stays off by default. Equality repair
retains its existing events, original audit and immutable cone-bit checks.

## Deliverables and limits

`rc9b.diff`, `diff_summary_rc9b.txt` and `static_checks_rc9b.txt` describe and
verify the fresh source copy. To recreate the copy and reports from this
folder, run `powershell -NoProfile -File .\prepare_rc9b.ps1` and then
`powershell -NoProfile -File .\verify_rc9b.ps1`. The preparation driver copies
the reference files and invokes the byte-preserving settings/native source
generators; it performs no Excel import or build.
`TEST_PLAN_rc9b.md` and `repro_rc9b.py` specify a
bounded future P02A_THIN E1/E0 test with explicit call/phase checks and optional
matched rc9a baseline comparison. The repro was written but not run. Runtime
coverage that does not occur in a short run must be reported as unexercised,
not silently inferred from zero events. Actual numerical acceptance and Fs
quality remain runtime validation items; source review cannot certify them.
