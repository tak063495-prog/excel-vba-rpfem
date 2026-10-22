# G1R20 Fix9 rc9c - test results (Winmax2, 2026-10-22 JST)

Change: rc9b + midnight-safe monotonic elapsed clock (new module RPX_ClockModule: RPX_Clock = GetTickCount64 on Win64, Timer fallback with day-wrap correction; default-off RPX_TestClockWrap hook, requires TEST_INJECTION=1). All elapsed-time Timer uses replaced; no other behaviour change. Codex-written; static checks PASS (731 non-ASCII lines byte-identical). Codex delivered RPX_Clock.bas (VB_Name RPX_ClockModule); renamed file to RPX_ClockModule.bas because the build names modules by file stem and module "RPX_Clock" collides with function RPX_Clock (build 1 compile FAIL). No code change.

Build: compile + reopen PASS, 64 standard modules, workbook sha256 d7f01647798d7232aedea703605d70776042355b0a96e7edb73fc9db9773aa21, DLL unchanged (66da1a21...).

P06N midnight (rc9b log + code review): time protection used QPC/Now-based clocks and stayed correct (HSD solve 173 stopped at 1489 s vs 1480 s limit across 00:00; solves 174/175 refused correctly). Only thin_layer_step seconds was wrong (-84046.3, true 2353.7). Reporting-only defect.

Repro (codex/repro_rc9c.py, P02A_THIN short: TARGET 320, 1 upper round cap 800, case limit 900 s):

| | W0 (hook off) | W1 (forced Timer wrap at 120 s) |
|---|---|---|
| run_end | PARTIAL_CERTIFIED | PARTIAL_CERTIFIED |
| Fs lower / upper | 1.302481156 / 1.313820965 | identical |
| wall | 332.5 s | 337.2 s |
| log elapsed_s | 5340 rows, 0..318.4, monotone | 5340 rows, 0..323.5, monotone |
| hook | inactive, wraps 0 | wraps 1, raw_jump -86398.8, corrected_step +1.21 s |
| thin_layer_step seconds | 54 | 55 |

The forced wrap was absorbed: there was no 86400 s jump, elapsed time was monotone, and the result was identical.

The driver's strict verdict was FAIL for both W0 and W1, for the same reasons. Neither is clock-related:
(a) pre-existing -1 "not applicable" sentinels (thin_layer_early_start remaining_seconds=-1 with remaining_source none; accuracy_dll_call packed_seconds=-1 x9). The test plan already lists these as failing strict acceptance without indicating a clock defect.
(b) The case-limit remaining-time stream was not observed (only 1 sample), so remaining-time continuity across the wrap is NOT verified by this repro.

Driver fix made during the run: the settings compare now treats same(None,None) as equal (Codex, 2 lines).
