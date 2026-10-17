# G1R20 P06 recipe and provenance handoff

Recorded 2026-10-16. This is a geometry-recipe replay with **newly solved fields and roots**, not a new generic cold mesh-generation success. Archived stress, velocity, Fs, solver vectors and root records were not used as numerical seeds. The current-session V1/V5 lower Fs was used only as an initial location hint; each imported geometry obtained its own new q8 field and signed root.

| Case | New lower | New upper | Symmetric width | Own q8 roots |
| --- | ---: | ---: | ---: | --- |
| P06A | 0.84413330078125 | 0.859615052748042 | 1.8173754281% | both complete |
| P06N | 0.6767771508826014 | 0.6902587260388777 | 1.9723805913% | both complete |

No physical tolerance or acceptance policy changed. P06A uses `equal_reduced_phi`; P06N uses its original Davis equivalent policy. Strength is freshly assembled at every Fs. These are floating-point audited certificates, not interval arithmetic proofs.

## Geometry boundary of the current reproduction

Read **only** `xy`, `tri`, `material_id` from each archive mesh, rebuild materials/body/tractions/constraints from the canonical case JSON, then verify original model ID, material area and first moments against the canonical initial model. Do not read an old field or an old Fs from the same archive folder.

| Case/side | Geometry source relative to `validation/accuracy_20261013/results` | Elements |
| --- | --- | ---: |
| A lower ancestor | native1000 import, then existing uniform red refinement | 1000 → 4000 |
| A upper | `a_flowflip/P06A/best_upper_mesh.npz` | 2904 |
| N lower | `n_lower_anglefan5966_close/P06N/best_lower_mesh.npz` | 5966 |
| N upper | `n_flow8739_cap_qp_helper_close/P06N/best_upper_mesh.npz` | 8739 |

N lower mesh ID: `0b1cffd134d60c92b20ec7480cbbbbfe27695c3a725f6f5d31ec5ca097587cc8`.
N upper mesh ID: `44c89fcad8854b4b5ef36f061999d3033613a1e5b7193b2e2ef4a8eb0a8b7fbc`.
N physical model ID: `08cc8e73e85408e7a82cbe79b5716ef007cb917d00e021872c40e510b5d17525`.

The old native1000 was an adaptive history, not the direct current TARGET1000 mesh. Its log records TARGET1536/start ratio .35 → 603 triangles, band relocation → 605, legacy band refinement → 785, staged red/green refinement → 1000. A/N material areas remain `[83.75, 11.6, 207.15]`.

The confirmed historical N lower geometry chain is native1000 → uniform4000 → stress_fan(theta .6, background0; 354 marks) → 5350 → star angular fan5422 → star5594 → star5966. The confirmed upper suffix is flow_metric_flip(passes5, 115 flips) on2914 → local kinematic6106 → directional one-edge8739 (2118 marked, 1318 split edges). The actual upper local/directional command theta and earlier2914 construction are not fully recovered. These marks depend on old fields: do not claim fresh end-to-end regeneration of this geometry chain. G1R20 can initially expose an explicitly named geometry-only recipe entry and separately validate automatic generation later.

## New lower solve

Entry: `G1R19_python26_geometry_replay_20261016/geometry_recipe_candidate.py`.
For N, initial location was the **current-session fresh V5** lower `.6703709008826013`, with span `.005`; neither old `.676...` nor any old lower stress was read as a hint. Use q8 original lower free-load formulation, existing affine elimination and original full-space restoration. Search uses bounded bisection, the existing raw physical gates, original dual evidence where available, normalized root tolerance `1e-4`, LP calls at most30s, case450s, one thread.

N completed own bracket `[.6767771508826014, .6768552758826014]`, signs `[+1,-1]`, width `7.8125e-5`. Accepted lower is the original physical endpoint, not the opposite numerical sign endpoint. This fresh solve took427.489s in its wrapper and passed a separate no-solver original q8 audit.

## New upper LP → proximal QP → fresh original dual

Entry: `G1R19_python26_upper_goal_recipe_20261016/upper_goal_recipe.py`.

1. Set trial `Fs = fresh_completed_lower * (2 + .019999)/(2 - .019999)`. `.019999` is scheduling margin only; final symmetric-gap gate remains `.02`. N initial Fs was `.6904487260388776`.
2. Solve the fresh original upper LP (affine elimination permitted, then full restore), call limit30s, max350 iterations, one thread. Save full original-layout x and original dual separately. Solver status is metadata.
3. Audit restored x against fresh **uncondensed** q8 original `(E,b,G,h,cost)`. Positive work normalization is permitted only when `h==0`, exactly one original b row is nonzero, and its work ratio is positive finite. Divide the complete x by that ratio; then repeat all original equation/cone and independent physics tests. No normalization changes a coefficient, load or strength.
4. If LP x fails, use the original constraints unchanged with `P=.01 I`, `q=-.01*center`, and an additional **original cost cap `.999998`**. In Clarabel form: `A=[E;-G;cost]`, rhs `[b;h;.999998]`, cones `[ZeroEq, original product SOC, Nonnegative1]`. Strict tolerances `1e-13`, iterative refinement30/abs+rel `1e-15`, equilibration OFF, max350, call90s. Use the restored LP x as the first center, then at most one same-Fs QP recenter with that QP x. QP's modified dual is never an original root guide.
5. Save a field only after the unchanged full original-q8 physical **and witness** gates pass. Its root remains false until independently closed. For a next close point use `U - .95e-4*max(1,U)`; at that Fs try original dual feasibility first, then fresh LP/QP if needed. The main upper total budget600s was reached with full physical upper `.6902587260388777` and root still false; wrapper607.700s includes in-flight finishing work. This was not reported as PASS at that time.

The three accepted N physical upper endpoints were `.6904487260388776`, `.6903537260388777`, `.6902587260388777`; all derive from the formula and new fields, not archived Fs. The first two needed LP + QP + same-Fs recenter. Intermediate full primal residual failures were retained and rejected.

### Original dual convention and final root closure

For the original primal `E*x=b`, `G*x+h in K`, objective `cost*x`, original dual is `(y,z)`, with

`cost + E.T*y - G.T*z = 0`, `z in K`, `dual_objective = -b*y - h*z`.

Original-dual helper variable is the concatenated original `(y,z)` and is recovered from **helper solution.x**. Helper solution.z is unrelated and may never be treated as an original guide. Require all finite, original raw stationarity `<=1e-7`, cone `<=1e-8`, upper inverse objective `>=1.000003`, exact case/mesh/model/policy/q/Fs binding. Helper status alone is never proof and no helper output publishes a primal bound.

An extra original-dual recenter60s was rejected (cone `2.3659e-7`). A direct SOC projection candidate was also rejected because stationarity increased to `6.5873e-6`.

The final one-shot helper used the latest physical `U-.999e-4*max(1,U)` = `.6901588260388777`. Same original stationarity/SOC/separation constraints; auxiliary `P=1e-8 I`, q0, equilibration ON, strict `1e-13`, refinement30/abs+rel `1e-15`, max350, one thread,60s. It returned a full original dual with cone `5.6613869e-13` but stationarity `1.7849142e-7`, so it was rejected.

The sole final correction (`G1R19_python26_dual_linear_repair_20261016/repair_original_dual.py`) held z **bitwise unchanged**, solved once

`E.T*dy = -(cost + E.T*y - G.T*z)`

with undamped LSMR `atol=btol=1e-15`, conlim `1e12`, maxiter200, maximum10s, and set `y += dy`. Actual linear time `.002188s`; no conic optimizer call. Full fresh original re-audit then passed: stationarity `8.9245876586e-8`, cone `5.6613869e-13`, objective `1.0000030000010307 >= 1.000003`. A cached-only original search combined that guide with the full physical upper at U and closed normalized width `9.99e-5`. An actual cache miss was guarded by a forbidden-optimizer exception. This endpoint guide is only a discrete root sign; the existing physical upper field remains the continuum bound.

## VBA/native transfer and first parity checkpoint

Before invoking the single DLL, freeze canonical input/model/policy/q/mesh/Fs and the literal unscaled original CSR program. The DLL accepts the generic LP/QP/helper problem; it does not decide the physical audit. Preserve scalar/context identity through DoEvents, cancellation and publish. Only update the public bound/state after independent original audit. Reject/fallback must leave the previous certificate and field intact.

An economical initial Python/VBA parity check is **assembly + one fixed-Fs original LP** on the same imported geometry, followed by comparison of raw original residuals, physical value, exact model/context identity and chosen LP→QP branch. Next compare the capped QP program (`Pdiag`, q, original rows and cap), work normalization eligibility, then the original-dual signs. Vector equality is not required for a nonunique solution, but all raw original gates and root identities are. Do not claim native full2% completion from an intermediate parity checkpoint.

Copy only geometry into a runtime recipe asset. The fresh Python fields and regression proofs belong to review evidence and must not become hidden runtime seeds. The per-component manifest stores actual original source hashes; a later aggregate manifest must never pretend all components shared a single recorded solver start.

## Evidence entry points

* A: `G1R19_python26_geometry_replay_20261016/results/P06A_COMPOSITE_20261016T122851Z`.
* N lower: `G1R19_python26_geometry_replay_20261016/results/known_lower_20261016T122636Z/P06N_known_lower5966`.
* N upper physical derivation: `G1R19_python26_upper_goal_recipe_20261016/results/upper_goal_20261016T123423Z/P06N_known_upper8739_goal`.
* N final helper: `G1R19_python26_terminal_final_20261016/results/final_guide_20261016T124646Z/P06N_upper8739_final_guide`.
* N repaired upper and root: `G1R19_python26_dual_linear_repair_20261016/results/P06N_repair_20261016T125042Z`.

No additional precision experiments are required for this handoff. Preserve the earlier cold, interrupted, budget-expired and rejected attempts separately; sum per-case process work honestly rather than reporting geometry replay time as cold total time.
