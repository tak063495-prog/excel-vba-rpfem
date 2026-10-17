# Issue28: new bounded Python conditioning experiment

Six new upper-LP solves, q=8 and Fs=1, on the recorded same-topology 214-element P02A / D05A geometry. Each solver was limited to20 seconds and1 thread; final six child processes exited normally in 9.275 seconds total. This duration is for a small diagnostic only.

| Case / route | status | original objective | work error | physical | original dual stat | original dual guide |
|---|---|---:|---:|---|---:|---|
| P02A_none | Solved | 2.58315896593 | 3.3307e-15 | True | 1.5848e-10 | True |
| P02A_affine | Solved | 2.5831589672 | 6.6613e-16 | True | 2.1551e-11 | True |
| D05A_none | AlmostSolved | 0.224377457847 | 0.002701 | False | 5.4001e-12 | False |
| D05A_affine | AlmostSolved | 0.115267666455 | 0.0052884 | False | 1.0049e-07 | False |
| D05A_normalized_objective | AlmostSolved | 2.58315387249 | 1.6672e-06 | False | 1.9475e-10 | True |
| D05A_velocity_units | AlmostSolved | 2.58315896739 | 6.6613e-16 | True | 6.1404e-08 | True |

All results, raw vectors, source/input hashes and process exit records are under `results/fixed_point_02`. `VERIFY_MICRO.py` reassembles the original programs and replays all six vectors with zero optimization calls; it matches both passes and intentional rejections.

## What is verified

P02A is the physical source of D05A: coordinates x1000, gamma /1000, cohesion/phi/psi unchanged. The recorded coarse mesh connectivity is identical. Cost grows1000x, strain rows shrink1000x, reference-work RHS shrinks1000x, and jump rows remain the same, up to recorded floating-point errors. Exact positive primal and dual maps preserve the original objective. No physical material/load parameter or gate was changed.

The ordinary original and affine D05A attempts fail the original work/physical gates. A pure objective normalization gives a valid original dual sign guide after multiplying its dual by QScale, but the primal still has work error1.6672e-6 and is rejected. The exact positive velocity-coordinate substitution plus a common positive multiplier per cone yields physical and raw original-dual passes with work error6.66e-16 and stationarity6.1404e-8. Solver status is AlmostSolved; the pass comes from original audits, not status.

At Fs=1 the physical objective is about2.583159, above the upper witness threshold. Therefore every `witness` flag is false. This is a fixed-point inverse-sign diagnostic; it establishes neither an upper Fs bound nor roots nor GAP2% completion.

Zero dual has raw stationarity1906.2283737862776, normalized stationarity1 and original objective0. A relative1e-7 tolerance still rejects it by seven orders of magnitude and its objective is below1.000003. Thus relaxing the raw stationarity threshold alone cannot rescue the issue's reported near-zero candidate. A dual solved with normalized cost must be converted to the original objective's units before original stationarity/objective gates. The intentionally unconverted diagnostic is stored as rejected.

## Limits and retained failure

The current Issue28 raw400-element HSD/native/helper vectors were unavailable. This214-element Clarabel experiment does not reproduce the reported exact1429.671 residual, identify a DLL defect, or verify a VBA fix. It supports conditioning/unit restoration as a direction to investigate. Product source and prior results are untouched.

`results/fixed_point_01` is retained with two wrapper errors: the new harness initially treated the unchanged core's third return as original dual, although that API returns reduced dual. The first harness is preserved as `reports/harness_initial_v1.py`. Only the new wrapper was corrected to explicitly restore original duals, followed by the complete second run. All12 attempted optimizer calls across both runs are disclosed; the numerical core and gates were unchanged.
