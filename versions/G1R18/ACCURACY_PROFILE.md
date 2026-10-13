# G1R18 accuracy profile and production verification

The G1R17 solver/DLL remains the production numerical engine. The G1R18 VBA
candidate adds optional, independently audited lower mesh candidates and a
mandatory quadrature check for an explicitly selected accuracy profile. The
reviewed Issue #8 seepage modules remain unchanged. Python results alone do not
certify the Excel implementation or authorize a successful-accuracy merge.

| Setting | Default | Accuracy experiment |
|---|---:|---|
| `QUADRATURE` | existing setting | 8 |
| `FS_TOL` | 0.0001 | 0.0001 |
| `GAP` | 0.01 | 0.01, measured as `(upper-lower)/midpoint` |
| `ADAPT_ACCURACY_CAP` | `TARGET` | 24000, up to 50000 |
| `ADAPT_ACCURACY_ROUNDS` | 4 | 8; existing bounded continuation |
| `ADAPT_LOWER_TARGET` | 0 | 4000; optional new uniform lower mesh |
| `ADAPT_LOWER_FAN_ROUNDS` | 0 | 1 initially; allowed 0–4 |
| `ADAPT_CERT_QUADRATURE` | 0 | 8; both saved sides must be independently recertified |

`ADAPT_LOWER_TARGET` creates a separate lower candidate from the original input
definition and original physical loads, using the existing mesh generator. The
actual element count must fit the accuracy cap. It does not replace or refine
the saved upper mesh. Water/body loads follow the original input path, including
the Issue #8 hydraulic calculation, then the ordinary reference-load transform.

An angular fan marks endpoints of exterior, traction-free `load` edges adjacent
to soil with exactly zero cohesion. It splits every triangle incident to those
vertices into three centroid children, including cohesive and rigid neighbors.
Every old node and edge remains. Each child receives the parent's material,
rigid ID and both physical body-load arrays. Exterior boundary records and
constraints remain unchanged. A parent map and a per-parent area check are
retained. The split is a candidate, not a certificate. The lower endpoint is
solved again on that candidate mesh and the existing original `RPX_FieldGate`
must pass before the independently saved lower bound can improve.

With `ADAPT_CERT_QUADRATURE=8`, final lower and upper are separately restored to
their immutable witness meshes. Prepared state, lower cache and root brackets
are reset. Each side runs a fresh q8 original-program endpoint search and passes
the original field audit. Only a valid, endpoint-matching root bracket within
the current `FS_TOL` is committed. Old q4 roots do not become q8 roots. A failed
side retains its old field, mesh, quadrature and root metadata, and the precision
profile remains incomplete. The target guard also checks current run, input
epoch, analysis policy, requested certificate q and both root-valid flags.

Regenerate UTF-8 source and CP932 imports with:

```text
python versions/G1R18/tools/build_adapt_candidate.py
```

In a separate copy of the G1R17 checked delivery workbook, replace every module
present in `versions/G1R18/src/vba` and add `RPX_AccuracyMesh` if absent. Keep the
existing native DLL beside the workbook. Compile before running. The serial
owned-Excel harness does this without changing macro security settings:

```text
python versions/G1R18/tests/test_adapt_excel.py --compile-only
python versions/G1R18/tests/test_adapt_excel.py --fan-geometry-only
python versions/G1R18/tests/test_adapt_excel.py --fan-only
python versions/G1R18/tests/test_adapt_excel.py --quadrature-only
```

The geometry fixture tests a zero-cohesion edge next to cohesive triangles,
loaded and constrained boundaries, distinct body-load components, an atomic
cap rejection, preserved old edges, and an unexposed triangle. Numerical fan
and quadrature tests remain small control-flow checks, not P06 certificates.

The immutable witness state now includes the lower stress-variable prefix,
load-variable index and reaction records. These are reconstructed and copied
with each mesh so restoring an older lower witness after a rejected larger
candidate cannot audit that field with the rejected candidate's variable layout.
The real Excel uniform/fan regression retained the upper mesh and exact payload
while improving the lower from 71 to 139 elements; both restored original field
gates and both own roots passed. Evidence is retained in
`results/adapt_excel_1791898363593743200/summary.json`.

For production, apply the unchanged validated P06A and P06N input records to
separate workbook copies, with their distinct associated/Davis-equivalent
policies and original phi/psi values. Run one owned Excel process at a time.
Use q8, final tolerance 0.0001 and the profile above; log all chosen settings and
input/source/DLL hashes. Export the independently retained lower/upper geometry
and physical fields, not the current display mesh. Restore and rerun the full
original field audit for each side, then verify both own q8 root brackets and
the symmetric gap <=1%. UNKNOWN remains UNKNOWN; a gap below 1% alone is not a
pass. Keep logs and failed-candidate records, compare hydraulic/body/face loads
against the original input model, and re-audit exports in Python. P06 production
success and the remaining required regressions are merge prerequisites.

The same isolated harness can execute each production case serially and export
the separate physical witness payloads in the existing `RP25M001` binary format:

```text
python versions/G1R18/tests/test_adapt_excel.py --production-case P06A
python versions/G1R18/tests/test_adapt_excel.py --production-case P06N
```

These commands use TARGET1000, cycles8, cap24000, eight continuation rounds, lower uniform target4000 and
one lower fan by default, with original policy and material values loaded from
the validated case JSON. Optional CLI arguments change those mesh budgets. The
summary records source/input/DLL hashes, actual settings, native definition
area/policy, root brackets and post-restoration field-gate results. A successfully
executed test harness is distinct from `precision_target_met=true`; a partial
numerical result is retained and explicitly reports `precision_target_met=false`.

Flow-metric flipping and cone normalization are not enabled by this profile.
They require separate original-model audit and production validation. Existing
native interfaces already accept the new conforming mesh and q8 assembly, so
these VBA geometry/control changes do not require a DLL modification.
