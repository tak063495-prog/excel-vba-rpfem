# Original upper-program Farkas guide candidate

This is an isolated variant of the G1R18 accuracy profile. Apply its overrides
after the G1R18 overrides. It retains the existing native DLL and does not alter
the source or the running workbook of a G1R18 production experiment. Source and
input hashes must identify the variant used by each numerical result.

The full native HSD program keeps every assembled equality row and primal
variable. Compilation applies equality row normalization and KKT permutation;
it does not apply Python-style equality/facial reduction. Equality duals map
back as `y_original = y_normalized / row_norm`. Cone coefficients and offsets
remain the original assembled values, including prepared strength refreshes.

The candidate retains exact equality coefficients and right-hand sides before
normalization, as well as the original objective vector. That payload is reset
with the prepared state and is copied/restored with the immutable lower cache;
the retained-memory estimate includes its arrays and identity. The model,
input epoch, load phase, quadrature, dimensions, current solve and candidate
context must match. Any affine/condensed program is excluded.

The new acceptance path is deliberately restricted to q8 upper Fs programs
with an exactly zero original objective and exactly zero physical cohesive and
fixed-load costs. A positive multiplier normalizes separation to 1.000003.
Compensated sums then audit the original, unscaled equation transpose and
original cone transpose. Acceptance requires stationarity <=1e-7, dual-cone
violation <=1e-8, positive separation >=1.000001 and normalization error <=1e-7.
Nonfinite vectors, coefficients, products or normalization are rejected.

`ORIGINAL_FARKAS_CAPTURE=1` optionally retains accepted original-program dual
proofs beside the diagnostic trace. The default is zero. Each immutable
`RG18F001` binary contains the original objective, equality matrix/right-hand
side, row norms, normalized and original equality duals, original cone matrices/
offsets and full cone dual. Its context records the exact mesh, material/load
identity, policy, phase and Fs/q. A ray capture contains no primal field or
continuum bound. The independent `tests/replay_original_guide.py` reader rebuilds
raw stationarity, dual-cone violation and separation from the saved matrices.
The native fixture also retains a fresh original program and physical field at
its negative endpoint, so a completed bracket can be replayed after the solver
has moved to another point.

An accepted ray has status `ORIGINAL_FARKAS_GUIDE`. It supplies only the
discrete positive sign for the upper root search. Its primal iterate and field
are erased and never captured as an upper continuum witness. The negative
side still requires the ordinary optimal physical velocity and original field
gate. A saved upper bracket may record an audited guide on its positive side,
with an optimal physical endpoint on its negative side. Other statuses and
quadratures receive no root promotion. Empty positive-side primal arrays are
allowed for this specific bracket; existing primal-field reuse then stays off.

Regenerate source and CP932 imports:

```text
python versions/G1R18_RayGate/tools/build_ray_candidate.py
python versions/G1R18/tests/test_adapt_excel.py --ray-only
python versions/G1R18/tests/test_adapt_excel.py --ray-adapt-only
```

Real Excel compilation, raw-gate boundary/identity tests, a native original
upper-program ray/negative-witness bracket, and recovery/old-witness retention
have passed on the small fixtures described below. These checks are control-flow
and original-model regressions; they do not certify P06A/P06N or the one-percent
goal for a production case. Python or G1R18 results are not results of this
separate variant.

The ray-only harness replaces the five ordinary G1R18 modules and then these
seven isolated overrides in a new owned workbook. It records both sets of source
hashes. Its unit programs check literal equality-row dual restoration, independent
stationarity and cone failures, nonfinite data, exact-zero costs, affine exclusion,
mesh/load/epoch/solve/quadrature identity, positive normalization, lack of ray primal
capture and unchanged physical bounds on rejection. The native test independently
requires an accepted original ray event and a q8 optimal physical negative endpoint
with a completed matching root. Run it serially after any production workbook
owned by this task has finished, unless a separate small owned test instance is
explicitly authorized. Parallel elapsed times are not a performance comparison.

`ADAPT_C0_WEDGE_DEPTH` defaults to zero. A positive depth enables one optional
upper candidate for a dry, gravity-only, single-region homogeneous soil with
exactly zero cohesion, no rigid bodies and the supported monotone six-vertex
slope polygon. It retains all six original exterior vertices and face
definitions and verifies the same area, material and body loads. Eight triangles
include a thin slope wedge. Unsupported geometry, positive cohesion, exterior
loading or nonfinite/nonpositive depth is rejected before changing the mesh.
The strength threshold only supplies a root-search hint. A new q8 upper endpoint,
original field gate and matching root within final tolerance are required before
the independently saved upper can improve. The lower field is not put into the
new upper payload. Existing lower geometry, field, value, quadrature and root
metadata remain immutable. Numerical failure restores both retained witnesses.

Real Excel evidence:

- `../G1R18/results/adapt_excel_1791900493885109200/summary.json`: original
  row-dual restoration and boundary/identity/nonfinite/affine gate tests,
  wedge geometry, and a native q8 positive original guide plus negative physical
  witness. Upper 1.50005035145014, own bracket
  [1.49992534722305, 1.50005035145014]; original guide stationarity at most
  5.04e-9, dual-cone violation zero, separation 1.000003.
- `../G1R18/results/adapt_excel_1791900646125391300/summary.json`: lower exact
  retention, independent upper adoption and injected numerical-failure rollback.
  Lower 26-element field 1.28674285650837 was exactly retained. Upper improved
  from a 26-element field 1.67993196508799 to the independently audited eight-
  element q8 field 1.50005035145014. Injected `CONIC_NOT_CONVERGED` kept both
  old fields/meshes/values and original field gates unchanged.
- `../G1R18/results/adapt_excel_1791901989652207900/summary.json`: final capture
  implementation and independent Python matrix replay. Three positive dual
  proofs retain original matrices and full restored duals. The negative physical
  field retains its freshly assembled original program. Replayed positive
  stationarity is 5.0399e-9 at the closest endpoint; negative raw equation error
  is 3.7690e-14, raw cone error and objective are zero. The own q8 bracket is
  [1.4999253472230507, 1.500050351450136].

Earlier failed fixtures are retained. A nonfinite-depth comparison was split
into separate finite/range checks because VBA boolean operators do not short
circuit. The native-root test calls the root getter before checking its ByRef
outputs; relying on those outputs within the same compound expression produced
a test failure. No physical tolerance or root requirement was weakened.
