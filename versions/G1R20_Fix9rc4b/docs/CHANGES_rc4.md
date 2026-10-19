# G1R20 Fix9 rc4: upper settings only after a taken early start

`vba_rc4/RPX_Adapt.bas` replaces only `vba/RPX_Adapt.bas` from the supplied
rc3b source set. Use it with the other 62 original modules in `vba/`.
All 63 supplied files are preserved; their SHA256 hashes are recorded in
`checks_rc4/baseline_sha256.json` and `static_checks_rc4.txt`.

## What changed and why

The thin chain can use a stronger upper marking theta and a larger upper cap
after `Fix9ThinEarlyStart` actually admits an early start in this analysis:

| New key | Default | Inclusive valid range | Conversion |
|---|---:|---:|---|
| `ADAPT_THIN_EARLY_UPPER_THETA` | 0.9 | 0.05..1 | `Fix8Number` |
| `ADAPT_THIN_EARLY_UPPER_CAP` | 7000 | 16..50000 | `CLng(Fix8Number(...))` |

These use the same bounds and validation as the corresponding base keys.
As in rc3b, `Fix8Number` rejects nonfinite or out-of-range values rather than
saturating them to a bound. Cap conversion retains the base key's `CLng`
rounding behavior, including ties to even; it does not introduce a stricter
integer-only setting. Negative values are invalid. Omitting the new keys
selects the defaults above, rather than disabling the overrides.

The module Boolean `fix9ThinEarlyStartTaken` resets to False at the first
executable statement of every `RPX_RunAdapt`, before settings can raise an
error and before any mode-specific path. Its only True assignment is next to
the successful `Fix9ThinEarlyStart = True` return. Per-candidate budget cleanup
does not reset it. A following adaptive run, including a non-Fs run, starts
False. Analyses that bypass the adaptive routine cannot call its private thin
driver; any later adaptive analysis resets the flag before using it.

`Fix8ThinLayerChains` retains the original base reads and validation in their
original order. Only after successful final thin detection, and only with the
flag True, does it replace upper theta/cap with the new values. Thus a declined
early start, mode 0, a successful pilot path, and ordinary later thin entry
retain the rc3b base values. All non-thin paths avoid the new setting reads.
The base keys remain subject to their original validation even on an early
route. Rounds, lower settings, flip passes, mid steps and all other chain
options retain their rc3b behavior.

In early mode 1, `Fix9ThinEarlyBudget` evaluates the proposed EARLY upper
theta/cap directly; the flag need not be True before admission. Its six
thresholds, comparison order, lower schedule, lower settings, deadline
validation and witness predicates are unchanged. Mode 2 still bypasses the
budget helper and deadline reads, and requires all original coarse witness
checks; a successful mode 2 start uses the EARLY upper settings in the driver.
Mode 0 remains silent at the early probe and uses the base chain settings.

The supplied P06N Python simulation motivates the defaults. The task reports
2.42% GAP at theta 0.8/cap 3500, and 2.008% at theta 0.8/cap 7000. The supplied
`docs/P06N_vcoarse_rc2_ut09_uc7000.log` records theta 0.9/cap 7000 reaching
GAP 0.01800556775832869 (1.80056%), with 5179 upper elements at upper round 9
and 4157 lower elements at lower round 3. That log configures 14 upper rounds,
but reaches the target by round 9, within the retained production budget of
10. Its logged 502-element starting lineages include the mid-line build;
the task's approximately 400-element coarse mesh is the admission denominator.
These are supplied simulation observations, not new rc4 VBA results. No P06A
result for the new settings is supplied or claimed here.

## Admission arithmetic

Assume the unchanged coarse q8, complete-witness, finite-bound and own-root
checks pass. The default optional deadline is 0 (inactive); an active deadline
must still have the unchanged minimum 3600 seconds remaining.

| Budget check | P02A_THIN | P06N / P06A production settings |
|---|---:|---:|
| Coarse elements | 355 | approximately 400; arithmetic uses 400 |
| Upper rounds, minimum 10 | 3 < 10: FAIL | 10 >= 10: pass |
| Lower rounds, minimum 3 | 2 < 3: FAIL | 3 >= 3: pass |
| Last scheduled lower round | 1+2-1 = 2 <= 3 | 2+3-1 = 4 <= 10 |
| Proposed early upper theta | 0.9 >= 0.8 | 0.9 >= 0.8 |
| Unchanged lower theta | 0.9 >= 0.9 | 0.9 >= 0.9 |
| Proposed early upper cap ratio | 7000/355 = 19.718310 >= 2 | 7000/400 = 17.5 >= 2 |
| Unchanged lower cap ratio | 3000/355 = 8.450704 >= 2 | 6000/400 = 15 >= 2 |
| Mode 1 decision | declined: `upper_rounds_short` | admitted: `chain_budget_sufficient` |

P02A_THIN still fails first at 3 < 10, even though its proposed early budget
uses the new upper defaults. Its flag remains False, so its actual later chain
uses the input's base upper theta 0.8 and upper cap 3000, with lower cap 3000.
It follows the rc3b route. The task's rc3b measured P02A_THIN GAP of 0.450% is
context only; no new numerical measurement was made.

Both P06 cases pass the configured budget under the stated assumptions. With
the production caps, both ratio tests pass for any positive coarse count up to
3000 (the lower 6000 cap is limiting), so small variation around 400 does not
change that decision. Admission does not guarantee that all configured rounds
will execute or that either case will reach its target.

## Diagnostics and reproducing rc3b

`thin_layer_begin` records the effective `upper_theta` and `upper_cap`, plus
`early_overrides=0/1`. The marker is 1 when the early branch was selected,
including when its values were explicitly set equal to the base values.
The mode 1 `thin_layer_early_start` budget text adds `early_upper_theta` and
`early_upper_cap`; its existing `upper_theta`/`upper_cap` fields also show the
proposed effective early values. Mode 2 retains its existing no-budget event,
and a subsequent `thin_layer_begin` shows its effective settings and marker.

To reproduce rc3b's route and numerical chain settings exactly, set BOTH new
keys equal to that input's corresponding base values. With production base
defaults, use:

```text
ADAPT_THIN_EARLY_UPPER_THETA = 0.8
ADAPT_THIN_EARLY_UPPER_CAP = 3500
```

For a base upper cap of 3000, use EARLY cap 3000 instead. Keep all other
settings, including early-start mode and admission thresholds, unchanged.
Matching both values restores both mode 1 budget inputs and executed upper
settings. The added diagnostic fields and internal Boolean remain present.
No negative sentinel or missing key is needed. Setting early-start mode 0
disables early entry but is a separate route choice from reproducing rc3b.

## Verification and risks

`diff_summary_rc4.txt` is a unified diff against `vba/RPX_Adapt.bas`.
The replacement retains the original header, `Option Explicit`, ASCII bytes
without a BOM, and CRLF line endings. Only four procedures, one module Boolean
declaration and two corrected explanatory comments differ. Accuracy rules
(q8, FS_TOL, audits, own roots, certificates), lower retries, all other
procedures and the other 62 modules are byte-identical to the supplied rc3b.

Reproduce the replacement/diff and run the retained static verifier from this
folder with:

```powershell
powershell -NoProfile -File .\checks_rc4\build_rc4.ps1
powershell -NoProfile -File .\checks_rc4\verify_rc4.ps1
```

`static_checks_rc4.txt` records source hashes, byte format, unchanged scope,
changed-procedure If/End If and For/Next balance, declarations and source-based
admission/selection checks. These checks inspect source and arithmetic; they
do not compile or execute VBA. No Excel, VBA, solver, network or git was used.

Higher upper theta/cap can increase mesh size, memory and solve time. Indicator
concentration, closure, stalls, target exits, failed steps, retries and deadlines
can still shorten the chain. The supplied P06N simulation does not prove rc4
VBA accuracy or runtime, and P06A with these settings remains unvalidated.
No accuracy threshold or certificate rule was relaxed to obtain admission.
