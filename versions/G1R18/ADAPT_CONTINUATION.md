# Adaptive accuracy candidate

The ordinary G1R17 adaptive route is retained. Intermediate Fs root searches use
0.001 instead of 0.005; final searches still use the configured `FS_TOL` (the
default is 0.0001) and requested quadrature.

| Setting | Default | Range and effect |
|---|---:|---|
| `ADAPT_ACCURACY_CAP` | `TARGET` | At least `TARGET`, at most 50000. A larger cap enables refinement after the ordinary route. |
| `ADAPT_ACCURACY_ROUNDS` | 4 | 0–20. Zero disables continuation. |

Continuation respects `ADAPT_MESH_POLICY`; `ROBUST_REFINE` and
`LEGACY_THEN_REFINE` use the existing conforming red-green refinement for cases
without rigid regions. Each attempt grows by at most a factor of 1.5 and stays
within the configured cap. Stagnation triggers a retry ranked by the retained
bound-gap indicator; two successive gains of at most 1e-6 stop the continuation.
Recoverable numerical or geometry failure restores the complete baseline. Memory
unavailability, cancellation, unexpected errors and restoration errors propagate.

Every evaluated candidate still uses `Evaluate`, `SaveState`, `RPX_FieldGate`,
the original input/model identity checks, and separate immutable lower and upper
witness meshes. Root validity, quadrature and target acceptance are not promoted.
No DLL change is needed for this VBA control and mesh-budget change.

`src/vba/RPX_Seep.bas` and `RPX_SeepUI.bas` are the reviewed Issue #8 source;
their CP932 imports are copied unchanged from the Issue #8 candidate.

## Real Excel verification

The serial Python harness needs an existing Windows Excel and permitted VBA
project access. It preserves automation and Trust Center settings, starts a new
owned Excel process, and creates separate test files.

```text
python versions/G1R18/tests/test_adapt_excel.py --compile-only
python versions/G1R18/tests/test_adapt_excel.py
```

Excel 16.0 compiled all three candidate modules and the test module. The small
associated slope test uses TARGET64, q2, CYCLES0, GAP1e-12; it is a control-flow
regression rather than a P06 accuracy result.

| Route | Elements lower/upper | Fs lower | Fs upper | Gap / midpoint |
|---|---:|---:|---:|---:|
| Continuation disabled | 67/67 | 1.36116470940989 | 1.50828188339831 | 10.254% |
| Cap256, two continuation rounds | 142/142 | 1.38841098742785 | 1.47205042299660 | 5.848% |

Both outputs keep `RFsSearchIncomplete=True` and do not accept the deliberately
unmet target. Both saved lower and upper fields were restored to their own
original witness meshes and passed `RPX_FieldGate` again. The logs show that
only the enabled run entered continuation, two rounds completed, and final Fs
tolerance remained 0.0001. Source/import text matches for all three modules.

Evidence: `results/adapt_excel_1791889112434570400/summary.json` (initial complete
real Excel run), `results/adapt_excel_1791889206667774900/summary.json` (additional
physical gate checks) and the latter folder's `log_verification.json`. The second
harness initially failed while decoding a UTF-16 log as UTF-8 after the numerical
runs and physical audits had passed. That failure record is retained; the
BOM-aware log parser subsequently verified those same retained logs, without
rerunning or replacing the numerical evidence. The second test workbook was not
saved after that harness exception and is not a delivery artifact.

The recovery path was subsequently checked by injecting `CONIC_NOT_CONVERGED`
at `accuracy_1` from the isolated test module. The recovered run preserved the
exact baseline bounds, both 67-element witness meshes, and the incomplete flag;
both original physical gates passed again. The failed trial was not recorded as
an accepted accuracy round. Evidence is
`results/adapt_excel_1791889597845674700/summary.json`; reproduce with
`python versions/G1R18/tests/test_adapt_excel.py --failure-only`.

Full q8/P06 runs, indicator-stagnation activation, cancellation/memory-failure
paths, and the maximum 50000-element cap have not been verified by this small
regression. Accuracy attainment must be established by the numerical experiments
and production-sized validation separately.
