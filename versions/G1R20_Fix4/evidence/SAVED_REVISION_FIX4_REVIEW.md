# Saved result source revision regression

Fix4 changes the physical load phase and therefore changes the main source identity to `RPX_MAIN_SOURCE_V1_FIX4`. Existing Fix3 output must be revoked on opening the upgraded workbook, even when its physical input cells are unchanged.

`results/saved_revision_positive_01` passed the three-session migration test. The test saved an explicit synthetic display pair 1.11 / 2.22 with the published Fix3 `RPX_MAIN_SOURCE_V1` source record, imported the final Fix4 source without reanalysis, saved and reopened with events enabled. The old XML proof was preserved through the source upgrade. Reopening cleared both display values and reported `未解析（入力変更）`; physical input hashes remained equal before upgrade, after upgrade and after reopen. All 64 reopened code modules matched the frozen source, compilation passed, and all owned Excel processes exited without dialogs or monitor errors.

The source-record/display pair is synthetic and makes no numerical soil certification claim. The test binds the final saved Fix4 delivery book, its 64 source modules and unchanged single DLL by SHA-256 and verifies that monitored baseline, delivery, source and test files remain unchanged.
