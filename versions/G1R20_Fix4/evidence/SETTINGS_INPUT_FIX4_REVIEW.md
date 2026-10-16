# Main settings input review

## Reproduced on the unmodified published Fix3 workbook

`results/settings_initial_negative` ran 20 public input/route/target cases plus duplicate and out-of-range cases in owned Excel PID 10636. All immutable baseline hashes were unchanged; no dialogs, monitor errors or conic optimization occurred.

- **C0_AUTO logical false enabled the route:** both Boolean False and text FALSE returned True from public `RPX_C0AutoRequested` for a used c=0 soil. Unsupported 2 / bad also enabled it. This could select a different mesh strategy from the explicitly disabled input.
- **Fractional certificate quadrature disabled the gate:** `ADAPT_CERT_QUADRATURE=0.4` was rounded to zero by CLng; public `RPX_TargetMet` returned True for an equal synthetic pair without any certificates. 1.6 and 7.6 were also silently rounded. This was a synthetic gate test, not a numerical certificate.
- **Duplicate settings silently selected the first value:** TARGET rows held 1536 and 4096; `RPX_ReadInputs` accepted both and used 1536.
- **Supported seismic setup and actual reader disagreed:** public `RPX_SeismicSetup` preserved a nonzero KH=0.2 at row 110, but the row-100 setting reader returned zero. Such edits also lay outside the old main observation/source footprint.

`results/settings_missing_negative` separately confirmed in owned Excel PID 13896 that SeismicSetup, when exactly one seismic key was missing, inserted two new rows, reset an existing nonzero magnitude or leftward direction, and left duplicate keys. All baseline files remained unchanged and owned Excel exited.

## Fix4 contracts

The setting reader, producers, source record and idle observation share rows 4 through 120 via `RPX_SETTING_LAST_ROW`. Any key-column cell below the supported table, including a formula that displays an empty string, is rejected rather than silently replaced by defaults; description-only rows in A:C remain allowed. Keys use case-insensitive trimmed identity and duplicate/error keys fail. Existing supplied fallback values are preserved when a key is absent, including diagnostic sentinel fallbacks.

`RPX_P4` validates explicit typed settings before use: binary logical Booleans map False/True to 0/1; TRUE/FALSE text and numeric 0/1 have explicit meaning; C0_AUTO also retains OFF/ON aliases. Non-switch numeric inputs reject Boolean/error/nonfinite values. Counts must be whole, quadrature must be exactly 0/2/4/8, and documented ranges/enums are enforced before CLng. Unknown extension keys retain their values, with Excel errors rejected. Missing legacy keys retain the existing caller-provided defaults.

The common standard GAP and fallback is 2%; explicit user GAP=1% is preserved. Seismic setup adds each missing key independently and preserves existing values/formulas and rows.

## Regression

`tests/run_settings_input_fix4.py` checks the full candidate code and document-module identity and compiles each owned Excel session. It covers valid/invalid types, explicit Boolean switch meaning, count and quadrature fractions, duplicate keys, row 110 read/ordinary edit/linked-formula invalidation, row 121 rejection, no-op/formula-preserving seismic setup, and the 2% defaults/presets/fallback with 1.5%/2.1% synthetic target decisions. Synthetic target/display checks perform no conic optimization and do not claim numerical certification.

The final saved Fix4 book and all 64 source modules are bound by SHA-256 in `results/settings_positive_01`: all 79 cases passed across four owned Excel sessions. This includes explicit MODE/GRAVITY errors and logical values, case/spacing normalization, and formula-preserving missing-key setup. All owned Excel processes exited, no dialogs or monitor errors occurred, and the monitored immutable baseline, product and test files were unchanged.

## Candidate regression found before release

`results/settings_extent_negative_01` exercised the first saved Fix4 candidate directly without VBE access. A supported synthetic main source record was preserved, then an out-of-range D121 key was introduced by linked-formula recalculation or by editing with events disabled and saving/reopening. Both paths retained the old display sentinel. Ordinary direct D121 edits correctly invalidated it. The test also observed that the initial outside-key Find check did not reject D121 in the public reader; its final assertion failed, while both display negative observations and all owned process exits were recorded. This failed test and its source/book snapshots are retained unchanged.

The final candidate replaces that check with one shared key-column extent validator using End(xlUp), called from the common Guard input extent policy. It covers observation, recalculation, source commit, save/reopen and public input reads. As with the main 1000-row tables, even an empty-result formula outside capacity is an explicit overflow. The final positive harness seeds a valid source before injecting that unsupported formula and verifies invalidation, rather than attempting to certify an already-invalid outside-table formula.

`results/settings_extent_positive_01` is retained as a failed test-harness run. Both display paths and the ordinary-edit control invalidated the old pair correctly, but the test invoked an expected unhandled VBA error directly through Application.Run. Excel opened the expected runtime error 5 / `SETTING_ROW_OUT_OF_RANGE row=121` dialog instead of exposing its text in the COM exception. Its raw failure, dialog record and original owned-exit=False flag remain unchanged. `POST_RUN_EXIT_CHECK.json` separately records the later read-only observation that every recorded owner PID was absent; no process was terminated by that check.

The corrected harness in `results/settings_extent_positive_02` passed all three owned sessions. Display invalidation still uses the final saved book without VBE access. A test-only catch function is added only to the disposable reopened test copy for the final public-reader assertion, returning error number 5 and the exact overflow token without a modal error. All 64 product modules in that copy matched the frozen source and compiled. The direct-edit control, linked-formula update and events-disabled save/reopen cleared the old pair; all owned processes exited with no dialogs or monitor errors, and monitored files were unchanged. The delivered workbook and product source were never modified by this harness correction.
