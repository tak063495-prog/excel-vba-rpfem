# Generic intermediate pilots with strict terminal roots

This isolated sibling copies v4, then adds only an optional new-point limit to intermediate searches and a strict terminal closure stage. All original SOCP matrices, constitutive rules, canonical inputs, quadrature and acceptance thresholds are unchanged. v1/v4 remain frozen and their unsuccessful attempts remain reportable.

An intermediate `search(..., max_trials=4)` retains every original-q8 physical field it obtains. It stops before starting a fifth new Fs point. Cached, independently audited fields do not spend that point count. The original LP retry/recovery order and the shared wall deadline still apply, so this is a point limit, not a promise of four solver calls. A pilot is explicitly `root_complete=False` unless the unchanged signed-bracket and normalized 1e-4 width condition was actually achieved.

After either retained side makes the **physical** symmetric pair width <=2%, the driver searches each incomplete side on that field's own mesh using unlimited ordinary root search (`max_trials=None`) within the remaining case budget. Both complete original roots are still mandatory for PASS. A narrow physical pair with a failed, stopped or expired terminal root remains PARTIAL/BUDGET_EXHAUSTED. No pilot width replaces the final root tolerance. The standalone `search` default remains unlimited, preserving old callers.

The rest of v4's policy is preserved for this comparison: local-coordinate P2/Bernstein transport followed by full original audits; independently refined lower stress mesh; 80% original dissipation plus 20% kinematic upper indicator; bounded bisection proposals; initial span .005; one thread; 30 seconds per solver attempt clipped by the 900-second shared process budget. q8, physical/raw-stress tolerance1e-7, original-dual cone tolerance1e-8 and stationarity1e-7, normalized root width1e-4 and strict 2% pair acceptance remain mandatory. Modified-QP duals are not used as original-program signs.

The planned run is an explicit continuation of the **new v1 attempt from this session**. It snapshots that predecessor's fields, meshes, certificate metadata and summary, then independently reaudits them. Its original ~906 seconds and the new elapsed Python time are added in `cold_cumulative_seconds`/`adopted_path_seconds`. The unadopted v4 comparison adds about587 seconds per case to `all_current_session_trial_seconds`, using an immutable timing-only report; its fields are not imported. Both sums are distinct from parallel batch wall time. This is reported as a separate continuation attempt; it is not a fresh cold-start success and does not import any 20261013/20261014 historical field, Fs result or root. No flow-metric flip is enabled here, allowing its later effect to be compared as a separate stage.

```powershell
python -X utf8 validation/G1R19_python26_v5_20261016/supervise_p06.py `
  --prior-run-root validation/G1R19_python26_20261016/results/p06_early_20261016_2028 `
  --output validation/G1R19_python26_v5_20261016/results/continuation_TIMESTAMP `
  --additional-attempt-report validation/G1R19_python26_v4_20261016/results/fresh_native1000_20261016T114554Z/FINAL_COMPARISON.json `
  --pilot-trials 4 --case-seconds 900 --solve-seconds 30
```

The supervisor owns the Windows venv redirector and numerical child by exact PID, birth time and command, records return codes, enforces child-first deadline cleanup and runs separately bounded independent original-q8 audit processes. The v4 stop-helper iterator/list error is fixed in this sibling and a regression exercises child-first order and stale-identity rejection. This new code does not edit the frozen v4 source.

Twenty-four lightweight tests pass. Added checks cover point limits, cached inputs, incomplete-field retention, naturally completed strict roots, unchanged unlimited search, invalid limits, two mandatory terminal closures, rejected incomplete terminal roots, and owned-process stop order. These are interpolation/control-flow tests, not new numerical accuracy certificates. Numerical results are published separately with source and input hashes and the independent audit.
