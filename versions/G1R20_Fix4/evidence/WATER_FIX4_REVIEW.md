# G1R20 Fix3 seepage and coupling review

The review used the published Fix3 workbook without replacing product modules.
The test workbook adds only a test module; baseline files were hash monitored.
Raw first-pass evidence: `results/water_fix3_negative_01/summary.json`.

## Confirmed defects

1. **P2 — an active seepage Stop request is ignored.** In Fix3,
   `RPX_SeepSolve` processes Windows messages with `DoEvents` but never checks
   `XCancel`; the source assertion also checks input equality only. A test-only
   Windows timer waits for the ordinary public seepage run to enter its forward
   solve, then invokes the unchanged public `RPX_Stop`. Stop was received at
   iteration 1, but the run continued to iteration 258, returned an empty success
   message, and published `K2=True` while `XCancel=True`.

   Fix4 checks cancellation at producer/iteration/commit boundaries, including
   inverse evaluations and LM retries. Cancellation raises error 18 with
   `SEEP_ANALYSIS_CANCELLED`, clears the in-memory convergence flag, and cannot
   publish a completed result. The numerical iteration and convergence tests are
   unchanged.

2. **P2 — an invalid water unit weight can be accepted with inconsistent
   producer and consumer physics.** Fix3 input validation ignores gamma-w. The
   prescribed-head producer silently replaces nonpositive gamma-w with 9.81,
   while the coupling consumer uses the entered value. Public mode-2 runs with
   `B4=0` and `B4=-1` both succeeded, wrote 19.62 kPa at the first node using
   gamma-w 9.81, and applied mean vertical body force -20 (the dry value) to a
   fully submerged unit-square model.

   Fix4 uses one water-unit-weight reader for validation, FEM and prescribed-head
   producers, and the coupling consumer. Explicit nonpositive, nonnumeric, or
   invalid numeric cells fail with `INVALID_WATER_UNIT_WEIGHT`; an absent or
   empty setting retains the existing 9.81 default. Valid input mathematics is
   unchanged.

## Initial observations excluded from the defect count

The first exploratory result-edit case replaced every head by the same constant.
The resulting vertical pressure gradient was unchanged, so it did not demonstrate
a numerical error. That observation is not evidence of another confirmed defect.

## Regression entry point

`tests/run_water_fix4_excel.py --execute --negative --output <new-directory>`
uses untouched Fix3 as the negative baseline. Omitting `--negative` loads the
complete current Fix4 source and workbook event source into a test copy.
The suite covers active forward/inverse Stop requests, cancelled mode 0/1/2
entry, cancelled result commit, invalid gamma-w, a valid changed gamma-w, and
the retained blank default. It never replaces physical formulas or acceptance
thresholds. These are small operational regressions, not a full 26-case or P06
production recalculation.

## Expanded input audit

The immutable Fix3 matrix at
`results/water_inputs_fix3_isolated_negative_02/summary.json` completed92 cases in100.6s.
All22 valid controls passed. The input-check button incorrectly accepted51
invalid cases;42 selected invalid cases also completed withK2=True. These counts
are test cases grouped by input contract, not51 independent code defects.
`evidence/WATER_INPUT_NEGATIVE.json` records the exact case names and raw hash.
There were no dialogs or monitor errors, the owned Excel process exited, and all
monitored published files remained unchanged.

Concrete examples include fractional IDs selecting another existing reference,
duplicate permeability entries replacing earlier values, arbitrary inverse guess
suffix text choosing a different initial value, an ignored201st table row,
negative beta yielding unsaturated conductivity above its saturated value,
explicit malformed settings silently falling back or rounding, and invalid
coupling/suction switches silently choosing another physical model.

Fix4 validates those values at both input checking and actual readers. Valid
negative/zero heads and negative flux, numeric text, anisotropy, small positive
conductivity and beta zero are preserved. Boundary circulation follows the same
explicit wrap contract as main faces. The input-source context carries
`RPX_SEEP_SOURCE_V1_FIX4`, requiring recomputation of old seep results; its pure
capacity checks and the main extent validator also reject stale-field reuse
after a tail edit.

Inverse maximum is now an actual maximum: too-small values produce a validation
error, and the final accepted-state forward solve is reserved within the budget
and included in the reported count. Equations and convergence tests are intact.

## Final verified results

`evidence/WATER_FINAL_RESULTS.json` indexes138 passing actual Excel cases on the
final delivery workbookSHA256
`b010ab89c4e70540277475d1aab025db2545a24fd6697ad87856998737e94be6`:
92 input contracts (22 valid controls and70 malformed inputs),14 cancellation/
gamma cases,29 consumer/migration cases and3 inverse-budget cases. All128
source/import paths match across those final suites and remain unchanged;
all owned Excel processes exited, with no dialogs or monitor errors.

The inverse counter is independent test-only entry instrumentation. Fix3 max6
performs7 forward calls but reports6; max12 performs13 but reports12. Fix4
performs/reports6 and12 respectively. An entered5 is rejected before any inverse
forward call. The physical equations and convergence gates are not replaced.

Two failed harness attempts remain in the raw evidence. One preserved invalid
settings between cases; the corrected fixture resets every advanced setting.
The other required literalFalse for an empty K2 and trusted K2 alone after a
cancelled writer. The corrected test checks the authoritative persistentPENDING
state and verifies actual consumer rejection. Product sources are unchanged
between each failed attempt and its corrected passing counterpart; details are
in`evidence/WATER_HARNESS_NOTES.json`.
