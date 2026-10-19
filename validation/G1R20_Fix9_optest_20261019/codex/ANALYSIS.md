# Fix8 thin-layer GAP analysis

This is a read-only review of the supplied Python emulation, JSONL results, prototype source, and time model. No Excel/VBA execution or simulation was performed. GAP is `2*(U-L)/(U+L)`; the target is at most 2%. Python results are evidence about mesh strategy, not validation of the VBA solver or its root/certificate contract.

## Baseline and limits

The `end` records in `sim/results/P06A_vba/log.jsonl` and `P06N_vba/log.jsonl` confirm:

| Run | Best lower Fs @ elements | Best upper Fs @ elements | GAP |
|---|---:|---:|---:|
| P06A_vba | 0.842940 @ 3373 | 0.858523 @ 3070 | 1.8317% |
| P06N_vba | 0.673535 @ 4007 | 0.690957 @ 3432 | 2.5536% |
| P06N_vba_lcap8000 | 0.675575 @ 6445 | 0.690957 @ 3432 | 2.2513% |
| P06N_vba_late_lt09 | 0.675586 @ 5982 | 0.689607 @ 4946 | 2.0540% |
| P06N_vba_lt07_u14 | 0.674639 @ 2615 | 0.689607 @ 4946 | 2.1943% |

P06A reaches the target when runtime constraints are omitted. P06N's default lower chain stops at round 3: the proposed 6255-element mesh exceeds the 4000-element cap. Its accepted 4007-element witness is possible because the emulation checks the cap before angular-fan refinement. The upper chain later rejects 4946 elements against its 3500 cap. Increasing the lower cap improves Fs but the tested Fix8 variants flatten near 0.6756; this is an observed trajectory, not a proven limiting bound. With U=0.689607, reaching 2% requires L about 0.67595. Cap changes or a smaller lower theta alone therefore do not resolve the mesh-strategy deficit.

## Why the separate mid-line matters

The prototype describes a one-element-thick weak band whose continuous upper velocity cannot express an internal slip line until transverse edges are split. Fix8 merges those mid-line edges with the first indicator refinement, then flips/fans and solves. In contrast, `--mid-separate` creates the mid-line mesh and solves its own root before any indicator refinement. This changes the field driving every subsequent selection, rather than merely adding the same edges in a different order.

`P06N_coarse_midsep/log.jsonl` confirms accepted lower roots of **0.666591 @ 631**, **0.674388 @ 1686**, and **0.676868 @ 3545** elements. The last value meets/exceeds the prototype README's reported **0.67670 @ 3528** using a similar mesh size. This is strong evidence that the initial mesh/field sequence drives the lower deficit. Baseline coarse lower refinement instead stops at 0.65478 @ 1341 after its next root fails. Delaying lower refinement and reducing theta help some runs, but do not reproduce the mid-separate result.

`protomatch` additionally uses an upper solve on the lower mesh at `1.08*lower Fs` for its kinematic score, and a fresh upper velocity on the split mesh before flips. These operations match `proto/run_pipe.py`; Fix8 instead locates the best-upper density onto the lower mesh and prolongs velocity before flips. The supplied protomatch upper sequences include P06N 345/.7787, 371/.8129, 417/.7348, 489/.7114, 581/.7111 and P06A 345/.9329, 372/.9385, 414/.8914, 488/.8727, 583/.8729. The task reports exact prototype upper-trajectory agreement. That exact comparison cannot be independently verified here: prototype `results/pipe/*/log.jsonl` files are not supplied. The source operation order and available emulation trajectory are consistent with that report.

## Root failures and incomplete snapshots

Lower `root_failed` records include 1241 elements (`P06N_coarse_lt07`), 2403 (`P06A_coarse`), 2834 (`P06N_coarse`), 3501 (`P06N_vba_late_lt07`), and 3516 (`P06N_vba_lt07`). The inspected 1241/3516 trials return Clarabel `NumericalError`, fail the physical/witness checks, and exhaust 20 solves without a closed accepted bracket. Fix8 stops that lineage immediately. A plausible objective near the root does not establish an acceptable witness. Conversely, some successful protomatch lower trials also have `NumericalError` status but pass the physical audit; status alone is not the acceptance rule. A bounded retry from the last valid lineage with reduced theta can avoid making one failed mesh terminal without weakening acceptance.

Three supplied logs have no `end` record and must be treated as incomplete:

| Snapshot | Latest recorded step | Current best pair / GAP |
|---|---|---|
| P06N_coarse_midsep | Upper 8, 1444 elements | L=.676868, U=.690636 / 2.0136% |
| P06N_coarse_protomatch | Upper 5, 581 elements | L=.676711, U=.711099 / 4.9556% |
| P06A_coarse_protomatch | Upper 5, 583 elements, non-improving lineage | L=.845255, U=.872678 / 3.1926% |

The midsep lower chain already failed at 7054 elements, but the upper can still improve. Its upper-8 record is newer than `sim_summary.txt`, which stops at upper 7 / 2.84%. No final target claim follows from these partial runs. The prototype README separately reports completed P06N pairs at 1.884% and 1.801%, with its own audit/root limitations.

## Runtime is an independent barrier

`sim/sim/vba_timemodel.py` estimates upper DLL time as `3.45*(cones/22741)^1.2` seconds and upper HSD time as `20+0.0037*cones`. The fixed 5-second DLL limit is crossed near 31k cones (about 960 P06N elements); the 300-second HSD limit near 76k cones (about 2350 elements on Winmax2). These are calibrated estimates, not P06 VBA measurements. The model assumes 13 solves per root step, linear VBA overhead, and an uncalibrated lower-HSD cost of four times DLL.

Accordingly, the as-is S0 model stops P06N_vba at upper 8 / 2536 elements, leaving 3.10%, and P06A_vba at upper 9 / 3070, leaving 2.28%. Scaling DLL time gives S1 estimates of 3.32 h / 2.55% and 2.72 h / 1.83%, respectively; S2 assumes no upper DLL failures and estimates 1.33 h and 1.07 h. These exclude reaching the starting VBA witness meshes and do not establish Fix9 timing. Separate mid-line roots, same-mesh indicator velocities, fresh flip velocities, bounded lower retries, and per-solve budget scaling address distinct barriers; their combined VBA result remains unmeasured.
