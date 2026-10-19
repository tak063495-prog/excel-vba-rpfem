# Task: G1R20 Fix9 candidate (VBA) - align the Fix8 thin-layer chains with the Python prototype

You are working in this folder only. Shell: always use `powershell -NoProfile` (a profile-loading shell fails in the sandbox).
There is no git here. Do NOT run Excel, VBA, or any long computation. Do NOT touch files outside this folder.

## Layout
- `vba/` : the Fix8 VBA sources (63 modules). EDIT THESE ONLY. Changed modules must stay pure ASCII (no Japanese in code/comments).
- `fix8_orig_vba/` : untouched Fix8 copy, for reference/diff. Do not edit.
- `sim/` : Python emulation of the Fix8 chains (`sim/sim/vba_chain.py`), results (`sim/results/*/log.jsonl`), summaries
  (`sim/sim_summary.txt`, `sim/sim_timemodel.txt`), roadmap draft `sim/README_JA.md`, time model `sim/sim/vba_timemodel.py`.
- `proto/` : the earlier Python prototype (`run_pipe.py`, `meshops.py`, `mesh_accuracy.py`, `plib.py`, ...) and `PROTO_README_JA.md`.
- `docs/` : Fix8 README (Japanese) and SOURCE_MANIFEST.json.

## Part 1 - analysis (write `ANALYSIS.md`, English, concise)
Read the sim results and explain what drives P06N/P06A GAP (target <= 2%). Key facts already found (verify against the logs):
- Fix8 as implemented (from VBA witness meshes): P06A 1.83%, P06N 2.55% (lower chain dies at its 4000-el cap; lower plateaus ~0.6756).
- `--mid-separate` (band mid-line applied as its own first step with its own root solve, then indicator refinement) fixes the lower:
  P06N_coarse_midsep L 631 .6666, 1686 .6744, 3545 .6769 (prototype best .67670@3528). `protomatch` (= mid-separate + lower indicator from
  an upper solve on the lower mesh at 1.08*Fs + fresh upper solve before flips) reproduces the prototype upper trajectory exactly.
- Lower root failures (Clarabel NumericalError) at 1241-3516 el end the lower chain under the Fix8 stop rule.
- VBA time model: DLL upper exceeds DIRECT_DLL_SECONDS=5 at ~31k cones (~960 el in P06N); HSD exceeds 300 s at ~76k cones (~2350 el, Winmax2).
Some protomatch/midsep runs are still running; analyze what is there.

## Part 2 - code: Fix9 candidate in `vba/`
All changes must be confined to the thin-layer chain (`RPX_Adapt.bas`: `Fix8ThinLayerChains`, `Fix8ChainStep`, `Fix8BuildThinUpper`,
`Fix8BuildThinLower` and new private helpers), plus at most narrowly guarded hooks in `RPX_Solve.bas` / `RPX_AccuracyNative.bas`.
Hard rules:
- Non-thin-layer paths must behave byte-for-byte like Fix8. Every new behaviour is reachable only while the thin-layer chain is active
  (`Fix8ThinMode` true and inside the chain); any new global flag/floor must be set only inside a chain step and always reset
  (including on every error path / `CandidateFailed` / `CandidateDone`).
- Do not change accuracy audits, root rules or certificate logic: `RPX_FsEndpoint`, `RPX_ActualRootBracket`, `SaveState`,
  `RPX_CertFromAdapt`, `RPX_Audit*`, `RecoverEquilibrium`, q8, FS_TOL. Witnesses are only accepted through the existing Fix8ChainStep path.
- Auxiliary solves used only to build an indicator/velocity (items 2 and 3) must never become witnesses, must not change RBestLower /
  RBestUpper / RCurrent / certificate / bracket state, and must restore all state exactly like Fix8ChainStep does
  (study its save/restore block). On any failure of such an auxiliary solve, fall back to the Fix8 behaviour and log it.
- Every new behaviour gets a setting read with `Fix8Number` (validated range) and a Fix9 default; log it in the `thin_layer_begin` event.
- Keep VBA 64-bit compatible, Option Explicit clean, no new modules unless unavoidable.

Changes:
1. **Band mid-line as a separate first step** (`ADAPT_THIN_MID_SEPARATE`, default 1). The first step of each chain (upper and lower)
   splits ONLY the transverse (mid-line) edges of the thin materials (`RPX_TransverseEdges`), no indicator edges, no fan, no flips,
   then does the normal own-mesh root solve and acceptance (result 1/2 advances the lineage). Later steps use withMid=False.
   This mid step does not count against the upper/lower round counts (or document what you chose). If there are no mid edges, skip it.
   With the setting 0 the Fix8 behaviour (mid merged into step 1) is unchanged.
2. **Lower indicator from an upper solve on the lower mesh** (`ADAPT_THIN_LOWER_OWN_UPPER`, default 1). In `Fix8BuildThinLower`,
   before splitting, solve the UPPER problem once on the lower lineage mesh at Fs = `ADAPT_THIN_LOWER_OWN_UPPER_FACTOR` (default 1.08)
   * lineage lower Fs, take its velocity field and use `RPX_ThinUpperScore` on it as `su` (no point location). Single solve, not a root.
   Prefer the DLL-direct path; it is only an indicator, so if the solve fails or exceeds its time limit, fall back to the Fix8 located
   best-upper density and log `thin_layer_aux` with reason.
3. **Fresh upper solve before flips** (`ADAPT_THIN_FLIP_FRESH`, default 1). In `Fix8BuildThinUpper`, after the split, solve the upper
   problem once on the refined mesh at the lineage upper Fs and use that velocity for `RPX_FlowMeshCandidate` instead of
   `RPX_ProlongUpperVelocity`; on failure fall back to prolongation and log it.
4. **Lower chain survives one root failure** (`ADAPT_THIN_LOWER_RETRIES`, default 1; `ADAPT_THIN_RETRY_THETA_FACTOR`, default 0.5).
   `Fix8ChainStep` must report why it returned 0 (e.g. cap / no_edges / geometry / root_failed / crossed / cancelled). When the LOWER
   step returns 0 for root_failed or evaluate failure (not cap, not cancel, not fatal errors that Fix8 re-raises), retry once from the
   same lineage with theta*factor; only if retries are exhausted end the lower chain. Retries count as steps.
5. **In-chain HSD budget scaling** (`ADAPT_THIN_HSD_BASE`=20, `ADAPT_THIN_HSD_PER_CONE`=0.0037, `ADAPT_THIN_HSD_FACTOR`=1.5,
   `ADAPT_THIN_HSD_CAP`=3600). While a chain step evaluates, the HSD solve budget is raised (never lowered) to
   min(cap, factor*(base + per_cone*cones)) where cones is the model cone count of that solve. Reuse/extend the existing
   `RPX_HsdBudgetFloor` mechanism in `RPX_Solve.bas` (OptimizeHSDCore) so the floor is computed per solve from the cone count only when a
   new chain-only flag is set; outside the chain the flag is off and Fix8 behaviour is identical. Log the applied budget.
6. **DLL-direct upper time scaling** (`ADAPT_THIN_DLL_FACTOR`=3, `ADAPT_THIN_DLL_REF_SECONDS`=3.45, `ADAPT_THIN_DLL_REF_CONES`=22741,
   `ADAPT_THIN_DLL_EXPONENT`=1.2, `ADAPT_THIN_DLL_CAP`=600). While a chain step solves the UPPER problem, the DIRECT_DLL_SECONDS limit
   used in `RPX_AccuracyNative.bas` becomes max(DIRECT_DLL_SECONDS, min(cap, factor*ref*(cones/ref_cones)^exponent)). Chain-only flag,
   reset afterwards. Lower keeps the existing limit.

## Deliverables (in this folder)
- Modified `vba/*.bas`.
- `CHANGES.md`: per change - files/functions touched, the exact guard proving non-thin paths are unchanged, new settings with defaults,
  new diagnostic events, and any deviation from this spec with the reason.
- `ANALYSIS.md` (Part 1).
- `diff_summary.txt`: list of changed modules and changed line ranges vs `fix8_orig_vba/` (you may produce it with a short PowerShell
  Compare-Object script; do not install anything).
Finish by writing a one-paragraph summary as your final message.
