# Fresh Python P06 policy v4 — 2% stopping target

This is an isolated copy of the frozen `G1R19_python26_20261016/python` pipeline. All 49 Python files and the canonical input manifest were copied first; `BASELINE_SOURCES.json` records their original hashes. The original v1/v2/v3 pipelines, historical witnesses and active VBA remain unchanged.

The new policy changes candidate selection and control flow, without changing the physical program or acceptance gates:

- Stop immediately after either independently retained side makes the complete pair meet 2%; do not finish an unnecessary second solve in the cycle.
- Transfer upper P2 velocity with the existing triangle-local barycentric difference form. Transfer lower Bernstein stress in triangle-local coordinates, copying unchanged cells exactly. Every transported field remains only a candidate until the full original q8 physical/witness audit passes. A source root is never transferred to its child mesh.
- Refine the retained lower field on its own mesh using the existing stress variation/yield-proximity indicator. Optional angular enrichment has composed parent ancestry and never substitutes the upper mesh for this independent field.
- For positive-friction upper models, mix 80% normalized original dissipative work with 20% normalized existing kinematic score. Verify the strain/jump work decomposition against the unmodified original objective. Phi=0 models keep the previous kinematic indicator. This is a mesh heuristic, not an error certificate.
- Propose bounded bisection points and a relative initial bracket span of .005 by default. This avoids deliberately targeting an objective inside the narrow numerical-sign dead zone. Signed brackets, original-dual separation, field gates and root width are unchanged. The standalone `search` function retains its previous default and adds only an optional keyword.

The strict conditions remain q8, original physical/raw-stress tolerance 1e-7, original-dual cone tolerance 1e-8 with stationarity tolerance 1e-7, own-mesh normalized root width `(hi-lo)/max(1,(hi+lo)/2) <= 1e-4`, and symmetric pair width `2*(U-L)/(U+L) <= .02`. Numerical status is metadata. A modified QP dual is not an original-program sign certificate.

## Current fresh profile

The P06A/N initial mesh comes from the current session's canonical native mesh-only export, target1000/actual1078 triangles. No saved numerical field or root is imported. A/N share coordinates and connectivity, but each case is rebuilt with its own canonical record and associated/Davis policy. Mesh/context/native binary/input-record/export-summary files are copied byte-for-byte and hashed before solving. This preparation uses the validated model builder; physical model IDs agree exactly with the current canonical P06 models.

Each case receives 900 seconds total, 30 seconds per solver attempt further limited by remaining time, one thread, 8 refinement cycles, and a 32000-element cap. Best endpoints are published on expiry, but an expired run cannot become PASS. The supervisor owns the venv redirector parent and exact numerical descendants, captures wait return codes, kills child before parent on its hard deadline, and performs a bounded independent original-q8 audit in a separate process. All source and input hashes are frozen. This is a fresh Python solve timing, **not an Excel end-to-end benchmark**. The external native mesh generation time was not recorded and is explicitly separate, rather than reported as zero.

Example (fresh output required):

```powershell
python -X utf8 validation/G1R19_python26_v4_20261016/supervise_p06.py `
  --native-mesh versions/G1R19_2pct/results/mesh_only_1792150651304041200 `
  --output validation/G1R19_python26_v4_20261016/results/fresh_native1000_TIMESTAMP `
  --case-seconds 900 --solve-seconds 30
```

`run_case.py` can also run one generic case. `--initial-mesh` imports geometry only. `--seed` is rejected without an explicit `--prior-run` whose scope identifies a stopped fresh pipeline run. Such a continuation snapshots all input fields/roots and records the cumulative previous-plus-new Python time; historical saved-proof runs are rejected. The current P06 profile does not use this option.

## Verification before the new solver runs

Sixteen light tests pass: previous shared-budget/2% boundary/root/expiry regressions, polynomial lower and upper transfer, immutable source fields, material/model rejection, original dissipation decomposition, phi0 fallback, and the early stop after a successful lower update. These control tests are not numerical accuracy evidence.

`preflight.py` starts no optimizer. It independently audited the new v1 P06 fields and their candidate transfers, and rebuilt the new native mesh with the canonical model. At that checkpoint, the c=0 band accounted for about 92% of the old kinematic score but no dissipative work; almost all work was in cohesive material0. The mixed score allocated about 81.5% there. Both old and local-coordinate transferred fields passed the original physical gate on the tested meshes. Results and exact source/input identities are under `results/preflight_20261016T114600Z/PREFLIGHT.json`; its saved candidate fields are **not bounds or root certificates**, and are not used by the fresh profile.

The preceding v1 trials ended BUDGET_EXHAUSTED at about 906 seconds per case: P06A gap4.1434%, P06N gap5.6127%. Those unsuccessful rows and all their fields remain intact. v4 is an additional improvement trial, not a replacement or an aggregation with historical strong pairs.
