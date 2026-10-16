# Seep producer / dependent soil state review

## Confirmed defect (P2)

Fix3 `RPX_SeepRun`, lines 994–1005, invalidated soil output only after a successful producer, and only if coupling was still ON at that time. The producer first revokes seep K2 while `RPX_InternalWrite=True` suppresses workbook input events. A legitimate coupling formula `=IF(浸透結果!K2,1,0)` can therefore switch 1→0 without invalidating a previously wet soil result. A failed producer left the old main display and `RPX_ResultCurrent=True` despite the consumer now requesting dry behavior.

The real Excel negative is saved under `results/seep_state_initial_negative`. It ran a real successful public mode2 seep producer, then seeded a clearly labeled synthetic main display/flags/source record (1.11 / 2.22; no numerical soil certificate). With K2/L2 unlocked on a protected seep output, public rerun revoked K2 but could not replace the rest of the protected output. Before→after: K2 True→False; coupling formula 1→0; current result / mesh flags remained True; the old main display survived. Excel PID 19932 exited; dialogs and monitor errors were zero; immutable Fix3 code/book hashes were unchanged.

The initial scratch harness also called `WaterSourceObserve` after the observation. That inherited diagnostic helper unconditionally queried a water head even though coupling had become OFF, returning error 11. This is not the asserted defect and is excluded from the formal regression runner. The entire initial output is retained without alteration.

## Fix4 implementation

- Capture coupling and prior `RPX_WaterApplied` before revoking seep K2. On every success/failure exit retire dependent soil results, even if coupling has since become OFF. Protected secondary cleanup cannot overwrite the producer's primary error.
- Preserve incoming Excel ScreenUpdating and EnableCancelKey; use xlErrorHandler during the producer so Esc can reach its cleanup handler. Preserve Busy / InternalWrite cleanup.
- Wire the shared strict water unit weight validator into the public input check, both mode builders, and existing error summaries. Invalid explicit values are highlighted and reported with `INVALID_WATER_UNIT_WEIGHT`; blank input retains the established default.
- Add cancellation checks around mode execution and final mode/source output, in coordination with Water/Seep/SeepWT checks.

## Regression scope

`tests/run_seep_state_fix4.py` uses real owned Excel sessions with full 63-module / ThisWorkbook code comparison and compile. Main output is an explicit synthetic display sentinel; this test performs no conic optimization. Its positive cases cover protected coupled output, formula-driven coupling 1→0, independent dry soil, healthy replacement, protected secondary soil output with an invalid-gamma primary error, and incoming application-state restoration. Numerical regression is separately performed by the release workflow.

`results/seep_state_positive_01` passed all five cases against the final saved Fix4 artifact/source binding. Every owned Excel process exited. Incoming ScreenUpdating, EnableCancelKey and WaterSuspend were restored; the synthetic result tests made no numerical soil certification claim. Monitored baseline, source, test and delivery files remained unchanged.
