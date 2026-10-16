# Inherited regression harness contract update

The first inherited final run stopped at `results/inherited_final_01/setup_control/summary.json` after all 18 setup/control cases executed. Seventeen passed. The remaining `run_invalid_mode` case already rejected the malformed MODE with error 5, source `RPX_P4`, description `INVALID_SETTING MODE`, and recorded the expected cleanup/state values. The inherited assertion still expected the later legacy reader's Japanese description. This was an outdated test expectation; no product change was made.

The original FAIL summary, runner snapshot, result values and ownership flags remain unchanged. The owned Excel process exited, no modal dialogs or monitor errors were recorded, the original workbook was unchanged, and all 63 product source files were unchanged. The recovered closed-window 1400 observation remains in the raw record as an explicitly tolerated monitor race.

- Original FAIL summary SHA-256: `e12d00a263a3e5e41179b6c10a3416d8c6b4bcc50710f5c69b66c9f64e1f68f3`
- Original subrunner SHA-256 recorded in that summary: `21036db4d11ed5446226769d140e1bf2a44379fe5c811ed2e2353a74d9f81cfa`
- Updated setup/control runner SHA-256: `972d9eaa184fae641697fd4334496960b3129b42f961b4d3e8e51e3ec375ead5`
- Updated inherited driver SHA-256: `43175b5b52771c9a43cd9158a88e80324d15a745fa2f2af60e06175ce2af0c88`

The Fix4-positive assertion now requires the exact error number/source/description triple `[5, RPX_P4, INVALID_SETTING MODE]`. The existing application state, Busy/InternalWrite, result-state and cancellation-key assertions remain in place. The published-only negative control retains its original legacy message expectation.

The inherited driver accepts `--only` with explicit group names, preserves the existing serial order, validates names, and records requested, selected and skipped groups plus `full_suite`. Omitting the option still selects all eight groups. The resumed final run selects seven groups (`setup_control`, `seep_ui`, `water_source`, `water_restart`, `saved_edit`, `saved_missingproof`, `saved_openchanged`), preserving the already-passed presets evidence from the first run. A filtered PASS describes the selected groups; completion of all groups requires combining the preserved presets run with the seven resumed groups.

Python compilation, command-line help and pure selection checks (all eight defaults, seven resumed names in standard order, duplicate suppression) passed without starting Excel. Both runners were frozen before the resumed Excel run. No product source, distributed workbook, historical raw evidence or Git state was changed by this harness update.

The remaining inherited assertions and their VBA probes were reviewed for the strict Fix4 contracts. No other contract mismatch was found: water boundary wrap fixtures are explicit, the inverse evaluation allowance of 12 exceeds the single-unknown minimum of 6, mode/setting fields in healthy fixtures are valid, accepted negative heads and formulas remain valid, and source-proof error tokens are unchanged. These are static audit conclusions; the resumed Excel run supplies the execution evidence.
