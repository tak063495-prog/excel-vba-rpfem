# rc9b short verification plan

The source-only delivery does not run Excel, VBA, a build or the repro. Static
results are in `static_checks_rc9b.txt`; numerical acceptance remains a later
runtime check. Use an rc9b workbook prepared from `vba_rc9b/`, the unchanged
DLL. `repro_rc9b.py` embeds the predecessor P02A_THIN input and configuration
helpers; no external harness/input file is needed at runtime. The script
copies the supplied workbook into its output directory.

## P02A_THIN, at most approximately 25 minutes total

Run later from this folder (the workbook path is required):

```powershell
python .\repro_rc9b.py --workbook C:\path\rc9b.xlsm --out run_rc9b --scenario all --wall-limit 1500 --rc9a-baseline C:\path\rc9a\repro_summary.json
```

Use a previously recorded rc9a scenario-0 baseline with the same P02A_THIN
short configuration and unchanged DLL. The predecessor baseline is
`../codex_fix9rc9a_repro/repro_rc9a.py`, not a new Excel run required by this
task. The default 1500-second combined wall limit covers E1/E0, startup,
mesh creation and close, with a 30-second cleanup reserve per scenario.
The first supervised work allowance is about 720 seconds; unused time can
be available to the second. Keep the
short 320-element target, one upper round and 800-element caps. Do not raise
solver budgets, relax audits or enable fault injection to obtain coverage.

| Scenario | ACCURACY_DLL_UPPER_EQUILIBRATE | Required behavior |
| --- | --- | --- |
| E1 | Preset/default 1 | Only literal current original upper primal calls receive the new override. |
| E0 | Explicit 0 | Disable the override and use the unchanged rc9a selector. |

An absent-row default of 1 is checked in source. Setting 0 means legacy
selection, not a forced packed zero: a thin original upper call can still
pack 1 through ADAPT_THIN_DLL_EQUILIBRATE=1. E1/E0 can consequently have the
same packed bit and numerical trajectory in this short case.

## Required log and result checks

1. Match `accuracy_dll_call` events by call ID, program role and dimensions.
   In E1 every original upper primal call reports upper setting 1 and actual
   packed equilibrate=1, including an eligible retry. Lower, helper and
   auxiliary calls must retain their existing selector. The helper must be
   labelled original_dual_helper and use ACCURACY_DLL_DUAL_EQUILIBRATE. In E0
   compare the actual bit against the legacy thin/generic/retry setting;
   never require unconditional zero. Check no new rejection path/event exists.
2. Require unique native call IDs and match vector audit/capture events to
   the candidate's origin. A failed retry may have a newer last_call_id than
   candidate_call_id; the restored vector must retain its first call's ID.
   An early native exit must show reset status, never a previous call's status.
3. Check each `accuracy_dual_vector` role and its attributed call. In
   `direct_dll_end`, legacy solver_status/raw_eq/raw_cone must match the saved
   primal phase. Separate dual fields must identify the audited dual call and
   role. A primal-only accepted trial must have dual_audited=False and zero
   unavailable dual residual/objective fields. A complete rejected dual audit
   may have dual_audited=True. No trial may inherit an earlier trial's values.
   Missing helper/retry/primal-only coverage is explicitly unexercised.
4. For original_primal acceptance, require the unchanged raw primal residual
   gates, status 1/4 where the thin upper restriction applies, the q=8 physical
   audit, measured value<=0.999999 and repeated commit/endpoint audit. A
   below-target DLL objective alone is insufficient. A passing ordinary
   original_primal trial must skip HSD for that same solve/Fs; helper-first
   work already performed can remain. Rejected candidates keep normal HSD
   fallback. The short case need not reproduce the seven historical dumps.
5. Compare verified finite lower/upper Fs and completion/audit status against
   the matched rc9a baseline, using --fs-tolerance (default 0.0001). E0 should
   reproduce rc9a within tolerance. E1 should give the same or better lower
   (no smaller beyond tolerance), upper (no larger beyond tolerance) and gap,
   with valid current certificates. Record wall/HSD time as observations,
   not a guaranteed speedup. Compare E1 with E0 separately as a paired check.

Keep `repro_summary.json`, each `scenario_E1/result.json` or
`scenario_E0/result.json`, worker state, settings and analysis/event logs.
Without the optional archived baseline, the independent rc9a comparison
remains UNVERIFIED; E0 legacy selection and the paired E1/E0 checks can still
pass. Required missing log evidence, invalid Fs or unexercised necessary
coverage must remain failed/incomplete, not pass.
Any timeout must terminate only the script-owned Excel instance and preserve
its diagnostics. A 25-minute run that cannot finish is a bounded test result;
it does not authorize extending this task or claiming a numerical regression
without examining its logs.
