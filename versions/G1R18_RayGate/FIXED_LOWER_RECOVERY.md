RPX_TryFixedLowerRecovery is an experimental same-Fs primal recovery helper.
The RayGate experiment now also has an optional Fs connection. Its default
LOWER_FIXED_RECOVERY=0 preserves ordinary Fs behavior. The ordinary G1R18
production profile and the running P06A baseline workbook have not changed.

The eligible input is a full, unreduced, reference-total, soil-only lower
program at q8 with raw physical auditing enabled, RFs active, and a
NUMERICAL_UNKNOWN or NUMERICAL_FAILURE status. Its original cost must be
exactly -lambda, with all other coefficients exactly zero. Model, input epoch,
quadrature, dimensions and compiled context must agree. Nested AF contexts,
other status, q, phase, cost, rigid bodies and nonfinite arguments are rejected.

The helper creates a fresh lower program at the same Fs, clears its objective,
and adds lambda = 1.000003. It uses the existing RPX_Affine exact-form singleton
and homogeneous two-variable substitutions, identically zero SOC-face
reduction and positive block/column scaling. A reduced HSD solve is allowed to
produce a restored primal candidate only. Reduced duals, infeasibility rays,
solver status and root flags do not establish a free-problem result.

After affine restoration, the helper assembles the original free-lambda lower
program again and independently runs RPX_OriginalPrimalGate and the original
RPX_FieldGate. All physical raw errors retain their existing 1e-7 limit. The
finite measured lambda must be within 1e-7 of 1.000003 and at least 1.000001.
The returned array is a primal witness at that Fs. It is not an optimum or a
root-completion claim, and it does not install a certificate or update bounds.

When the optional Fs connection is enabled, a known numerical failure first
offers the current original iterate normalized by one positive multiplier to
lambda=1.000003. That candidate is accepted only after fresh original free-load
equations/cones and the full q8 physical/raw gate pass. Invalid/nonfinite
normalization is rejected. If that original primal is unavailable, the fixed
load affine HSD solve described above is attempted. The diagnostic route is
ORIGINAL_ITERATE_PRIMAL or FIXED_LOAD_AFFINE_HSD. A fixed-problem failure never
becomes a sign or field. LOWER_FIXED_RECOVERY_SECONDS defaults to 120 and must
be positive and at most 600. Physical-gate, programming and model errors are
not made recoverable by this setting.

An accepted same-Fs field receives ORIGINAL_FIXED_LOWER_WITNESS, never OPTIMAL.
It supplies the positive lower endpoint only. An independent negative sign
requires a full, unreduced, original FREE-lambda dual with exactly -1 at the
load variable and zero cost elsewhere. The row dual is restored as
y_original=y_normalized/row_norm. Original compensated
cost+E.T*y-G.T*z stationarity must be <=1e-7 and original dual-cone violation
<=1e-8. Its objective implies lambda_cap<=1.000001-2e-6. A fixed or affine
dual, altered cost, nonfinite data, wrong q/model/input/phase/layout and other
objective scales are rejected. The dual receives ORIGINAL_LOWER_DUAL_GUIDE
and never supplies a physical field or continuum bound. Optional capture uses
LG18D001 with full original matrices and vectors, replayed independently by
tests/replay_original_guide.py.

A recovered root needs that separate original negative dual, q8, matching
model/input/phase/policy and endpoint identity, correct signs and ordered width
<=1e-4*Max(1,midpoint). Old evidence from another quadrature is not promoted.
The saved lower and upper witnesses/bounds are never changed by the recovery
helper or the trial alone. The usual full endpoint audit and owning adaptive
state are responsible for any later physical-bound publication.

Every attempt restores the original compiled equations, cones, cost,
permutation/KKT pattern, original raw snapshot, lower stress/load/reaction
layout, iterate/dual/slack, objective/audit/status, and original HSD proof
vector, HSD observation mode/identity and private capture vectors/flags. Numeric
factors are released and symbolic data rebuilt. The private
immutable lower model cache is not touched. Both original saved witnesses and
their bounds remain unchanged. Cancellation propagates after restoration.
Unknown numerical solves return no field; programming/model failures propagate.

HSD has an optional recoverySeconds argument (default zero, preserving the
existing HSD_SOLVE_SECONDS setting) to bound an isolated recovery without
editing the captured input sheet. The recovery API accepts 0 < seconds <= 600.
The DLL and ABI 0x00170005 are unchanged.

Real Excel compile, failure and physical checks passed in
adapt_excel_1791905387381788500 and again with HSD observation restoration in
adapt_excel_1791906288298223200. A cohesive 20-element model retained a valid
cache; a c0 24-element model retained its original cache-ineligible state.
Both returned lambda 1.000003, retained their original physical lower/upper
fields, and passed physical audits of those retained fields after rollback.
Before-solve and after-solve audit fault injection preserved both witnesses,
compiled layout, cache and UNKNOWN status. q4, non-UNKNOWN and nested AF calls
were rejected. No root was created. Independent replay of the literal free
program gave:

| Model | Raw equality | Raw cone | Free objective |
| --- | ---: | ---: | ---: |
| Cohesive | 1.667853355424853e-11 | 0 | -1.000003 |
| c0 | 2.877698079828406e-13 | 0 | -1.000003 |

The initial AHO3 recovery in adapt_excel_1791904907321458900 failed its strict
original linear-residual gate and was not adopted. The existing AF HSD
primal-only gate succeeded. The intermediate c0 test expected a cache that its
existing feature contract excludes; its failure is retained in
adapt_excel_1791905275431927900 and the final test checks the correct contract.
No numerical or physical tolerance was relaxed to obtain a PASS.

The existing original-dual guide / wedge / native own-root regressions also
passed with these passive overrides in adapt_excel_1791905485505122400 and,
with the optional Fs connection disabled, adapt_excel_1791908375927363000.
Additional q/epoch/model/policy/endpoint/width/missing-original-dual root
metadata rejection tests and full replay passed in
adapt_excel_1791908588145621600.
The final finite working-cost rejection and full replay also pass in
adapt_excel_1791910801411367300. A live model mutation invalidates the existing
bracket through RPX_ModelVersion; restoring input values does not resurrect a
stale root. The metadata tests perturb a copy of the stored candidate and
restore the untouched valid proof after rejecting the corrupted copy.
The 8-triangle q8 upper bracket remains
[1.4999253472230507, 1.500050351450136], with full original primal/dual replay.

```powershell
python versions/G1R18_RayGate/tools/build_ray_candidate.py
python -X utf8 versions/G1R18/tests/test_adapt_excel.py --fixed-recovery-only
python -X utf8 versions/G1R18/tests/test_adapt_excel.py --fixed-fs-only
python -X utf8 versions/G1R18/tests/test_adapt_excel.py --ray-only
```

The standalone fixed-load AF HSD tests passed again with exact private HSD
capture vectors/flags, original matrices/layout/cache and both saved physical
witnesses retained, including injected HSD_LINEAR_QUALITY_FAILED, in
adapt_excel_1791908292891767200. The optional Fs connection passed in
adapt_excel_1791908184243314500: recovered original primal at
Fs=1.37329120594012 (lambda=1.000003, raw equality 5.364619859449249e-11,
raw cone zero) and separate original negative dual at
Fs=1.373382541241769 (raw stationarity 1.5799121891516624e-10,
raw dual-cone zero, lambda cap 0.9999963471686844) closed the own q8 root.
Both saved bounds and fields remained unchanged. The literal matrices and
vectors passed independent fsum/hypot replay. Fixed-problem duals were unused.

The near-root AF fixed HSD attempt in adapt_excel_1791907938957103600 failed
HSD_LINEAR_QUALITY_FAILED and was not adopted. This known numerical failure
now returns False after full rollback. The near-root success above used the
freshly audited normalized original iterate; it does not establish successful
AF HSD solving at that point. P06 numerical recovery and production accuracy
remain unverified by these small fixtures. Keep the setting disabled for
production until a real P06 failure/recovery and full-state audit are verified.
