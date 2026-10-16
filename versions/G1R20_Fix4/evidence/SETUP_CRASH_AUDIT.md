# Setup after actual Fs: investigation record

One baseline session in `results/seismic_negative_/1792183690937194400` ended during the public mesh fixture after a successful actual Fs solve. A child Excel recovery process was verified by parent/process metadata and handled by its owner. That test had not logged the last internal VBA stage, so the failure cannot be assigned to Setup, SeismicSetup, mesh creation, DLL code or Office drawing.

The follow-up used the same unmodified Fix3 product modules and copied single DLL. Only a test helper added persistent boundaries before/after Setup, SeismicSetup, input table clearing and CreateMesh. Three cycles (9 public calls) passed in `results/crash_auto_01`.

A second owned Excel session repeated the same sequence with file tracing disabled for five cycles (15 public calls). It passed in `results/crash_auto_quiet_01`. The observed native report was DLL_READY with 70 native numeric calls, 70 NT bulk calls, 140 refinement steps and 150 input comparison calls. No recovery children were detected. Both owned processes exited; baseline book/source hashes remained unchanged; no dialogs or monitor errors occurred.

These eight actual-Fs / repeated-Setup cycles did not reproduce the crash. The original anomaly is retained; no root cause or DLL fix is claimed. A separate final saved Fix4 workbook regression directly calls public TestFs → Setup → SeismicSetup without any VBE access or injected helper to check the distributed artifact independently.

`results/pure_public_positive_01` passed that final saved-artifact check: three cycles / nine direct public calls, with actual Fs lower=9.99998978478745 and upper=10.0000099936856 in every cycle. The unchanged single DLL reported DLL_READY with 42 native numeric calls and 84 refinement steps. The owned Excel process exited, dialogs and monitor errors were empty, and all monitored files were unchanged. This adds an artifact-level regression; it does not establish a cause for the original isolated crash.
