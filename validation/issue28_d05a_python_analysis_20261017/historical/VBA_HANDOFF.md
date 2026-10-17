# Exact upper LP velocity units for D05A

This is a fresh Python validation policy, not an Excel timing result. The
length-times-1000 coordinates, original materials, loads, q=8 assembly and
physical-model identity were retained. No saved field or historical Fs was a
seed. Source and input files remain frozen.

Fresh result: lower 1.3028266247689342, upper 1.3254379799492289, symmetric gap
1.7206300415645823%, both 1622 elements. The first pair meeting the 2% policy
stopped refinement. Case wall budget elapsed 327.3986777999962 seconds.
Both own roots and both full original q8 physical witnesses passed an
independent fresh-process audit.

Evidence: `results/D05A_scalar_upper_20261016T111421Z/INDEPENDENT_AUDIT.json`,
`RESULTS.csv`, `COMPLETION.json`, `OWNERSHIP.json`, and standard side
certificates under its `cases/D05A/`. All 65 scaled LP attempts retained their
unchanged original-program digest. The raw numerical process OS exit code was
not captured by the detached launcher and is explicitly null; the independent
audit command exited 0. These are different claims.

## Transformation

Original physical LP:

`min c.T*x`, `E*x=b`, `G*x+h in K`.

Use `x = velocityScale*u`, with finite positive `velocityScale`. Let `L` be
positive diagonal equality-row normalization and let `R` contain one common
positive multiplier for every full original cone block. Solve

`q_new = velocityScale*c`

`A_new = [L*E; -R*G]`

`rhs_new = [L*b; R*h] / velocityScale`.

The cone list is the original zero equality block followed by the original
nonnegative/SOC blocks. Restore the full physical vectors as

`x = velocityScale*u`

`y_original = L*y_solver / velocityScale`

`z_original = R*z_solver / velocityScale`.

These are original LP duals after an exact coordinate/row transformation.
There is no modified QP or cap/floor dual in this path.

Python's existing upper assembly already normalizes equality rows, so its
`L=I`. Its only nonzero RHS is the positive reference-work row; the implementation
chooses `velocityScale=b_work` (D05A coarse: 7.874940184445002e-5).
VBA's `RPX_AccuracyExport` copies literal **raw** equality values/RHS, so it must
first use `L_i=1/||E_i||2` or the existing `XEVal/XERhs/XEqRowNorm` mapping.
Using only the raw VBA work RHS, which is 1, would not reproduce the Python
conditioning.

## Adoption requirements

- Original physical upper LP only: q=8, exact original layout/cost, original
  Fs/model/load-phase/input epoch and finite copied arrays. P=0 is required;
  a centered QP would additionally need P_new=velocityScale^2*P.
- Preserve raw E/b/G/h/c. Scale solver copies only. Equality multipliers are
  positive; each SOC head and all its tails use exactly the same positive R.
- A single positive reference-work RHS selects this automatic scale. Invalid,
  nonfinite, empty or unsupported layouts use the established fallback.
- Restore full vectors and run the unchanged original stationarity/cone/
  separation audit before using a dual sign. Raw tolerances remain 1e-7 for
  stationarity and 1e-8 for the dual cone. Solver status is metadata only.
- An upper bound additionally requires the unchanged original q8 physical
  primal and field gates, including work, cone and objective. Root completion
  still needs its own exact-mesh signed endpoints and the existing relative
  1e-4 width. No endpoint is inferred from a scaled-space residual.
- Positive work normalization is permitted only with zero original cone
  offsets and one positive work RHS, and only after full original re-audit.
  Existing phi=0 auxiliary completion is optional; D05A has nonzero phi, so
  that completion does not itself cure its conditioning problem.

The prototype records strict 1e-13 solver tolerances and refinement 30 / 1e-15.
When the restored optimization dual fails its raw gate, the existing original
dual-feasibility helper can provide an independently audited sign. Calling this
helper after the first useful physical LP avoids repeating identical strict LP
attempts; the frozen v3 driver retains its original four-attempt policy.

## Reproduction and implementation references

`run_case.py`: `scaled_upper_solve`, `normalize_upper_work`, `program_digest`.
The 49 copied support Python files are byte-identical to the bounded v1 tree.

```text
python -X utf8 validation/G1R19_python26_v3_20261016/run_case.py --case D05A --output NEW --target-gap .02 --case-seconds 900 --solve-seconds 30 --cycles 8 --max-elements 32000 --fan-lower
```

`results/scalar_probe_20261016T111340Z/PROBE_SUMMARY.json` records fresh fixed-Fs
diagnostics at 1.3 and 2.2. The run also checked random-vector equality,
objective and restored-stationarity identities and before/after original
program digests. The fixed-Fs fields were not seeds for the fresh cold run.
