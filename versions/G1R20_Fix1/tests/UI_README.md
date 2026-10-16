# UI preset regression

`UI_test_contract.py` checks the exact UTF-8/CP932 pair, the authorized three-routine
diff, Busy-before-write ordering, original-error propagation and unchanged preset
values. It does not replace the actual Excel tests.

After coordinating the single owned Excel slot, run:

```powershell
D:\Work\rpfem-venv\Scripts\python.exe -X utf8 versions/G1R20_Fix1/tests/UI_run_excel.py --execute --output <new-result-folder>
```

The harness imports the complete frozen 63-module source set into disposable
copies. It performs 24 state-machine cases: each of Accuracy, Speed2pct and G20
under a protected setting sheet, a missing GAP after earlier writes, Busy, and
normal success, both with ordinary and nested InternalWrite scopes. The late GAP
failure checks the original number/source/message and proves an earlier MODE
write occurred. Busy checks require unchanged settings. Nested scopes must stay
nested; ordinary completion/failure must invalidate results and leave subsequent
invalidation effective. Current-state flags are explicit test sentinels, never
fabricated fields or numerical certificates.

For migration it opens the actual published G1R19 workbook afresh for each of
the three entry points, imports all 63 modules, asserts that the seven G20 keys
are initially absent, and calls the entry without Setup. All seven keys and the
expected selected values must then exist. The exact physical input values,
formulas and analysis policy must be unchanged throughout every test.

The tests invoke no solver and publish no field or root. The harness records the
PID, birth time, process command, source snapshots, source-to-loaded code keys,
source/template hashes, modal dialogs and owned-process exit. A 180-second hard
guard can terminate only the exact owned PID with matching birth time. A final
PASS also requires process exit and unchanged original workbooks and sources.

`UI_negative_control_20261017.json` and `UI_negative_control_probe.bas` preserve
the earlier read-only review of the unmodified published G1R20. Both legacy
presets left InternalWrite true after a protected-sheet error; the G20 entry
failed after deleting only the new seven keys. Those simplified deletion results
are kept separate from the stronger actual G1R19 full-source migration test.
