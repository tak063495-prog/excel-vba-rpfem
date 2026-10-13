The current native DLL performs numeric kernels on VBA-owned arrays and
patterns. It does not independently assemble the physical model. The shipped
G1R17 DLL ABI is 0x00170005; its SHA256 is
72aaf9bfccc0acc186eb69cdcbabd93b27f7ce245ceac574193ba6602c675323.
The new primal-import and fixed-load recovery tests use this unchanged DLL.

Existing controls and limitations, read from the local source:

| Mechanism | Source | Control / limitation |
| --- | --- | --- |
| Equality normalization | RPX_Conic.RPX_Compile | Each original equality is divided by its stable row norm; XEqRowNorm preserves the inverse dual map. Original literal coefficients/RHS/cost are retained by RayGate. |
| KKT preconditioning | RPX_Solve.FactorSystem / HBuildMatrix | Permutation and diagonal scaling operate on the numeric KKT system; original residuals remain separately checked. |
| HSD backends | RPX_Solve.OptimizeHSDCore | HSD_LINEAR_BACKEND accepts AUTO_PIVOT, PIVOT_ONLY or LDL_CHECKED. Projective normalization and original-operator diagnostics are independent switches. |
| HSD budget | RPX_Solve.OptimizeHSDCore | HSD_SOLVE_SECONDS is checked at iteration boundaries. The native P06A 300-second attempt reached 341.74 seconds before rejection. New recoverySeconds overrides only an isolated helper; default zero preserves the setting. |
| Regularization / outer iterations | RPX_Solve | HSD regularization 1e-10 and 150 iterations are source constants. Several DD operator, matrix and diagnostic paths use the same regularization and must remain synchronized. |
| Linear refinement | native/rpfem_hlinear.cpp | Original/scaled residual target 1e-11, up to eight corrections with explicit stagnation guards. These controls are source constants, not an exposed run-time ABI setting. |
| Krylov correction | native/rpfem_krylov.cpp | Bounded restarted correction with existing factor/operator, residual target 1e-11 and epoch/finite checks. |
| Ordering / compile | RPX_Native, RPX_Conic, native/rpfem_compile.cpp | VBA passes full arrays and native contexts with explicit dimensions, patterns and permutations. Optional native ordering retains a VBA fallback. |

The existing RPX_Affine implementation already performs constructive singleton
and homogeneous two-variable substitutions, keeps one coefficient per original
variable plus affine offset, converts identically-zero SOC heads to tail
equalities, and applies positive Ruiz block/column scaling. It retains original
rows/cones and restores primal variables through AFMap/AFWeight/AFOffset. It is
restricted to a fixed-load zero-objective witness problem. Removed equality
rows have only textual history; a general original dual-restoration graph is
not implemented. Its use for free-lambda minimization would require objective
transformation, affine objective constants, structural refresh and cache maps,
and full original dual restoration. Reduced duals are currently rejected by
RayGate. This restriction must remain in place.

The implemented first recovery therefore fixes lambda = 1.000003 at the same
Fs, invokes the existing fixed/zero-objective AF pipeline, and tests the
restored primal on a freshly assembled original free-lambda program. This
requires no new DLL ABI and avoids treating a reduced fixed-problem dual as an
original root guide. Small cohesive and c0 tests pass; P06 recovery remains to
be established.

A further full-dimensional preconditioner could use positive diagonal column
scaling D, positive equality scaling R and one positive scalar C per entire
SOC or linear block. For original min c'x with Ex=b and Gx+h in K, solve with
x=D u, E_s=R E D, b_s=R b, G_s=C G D, h_s=C h and c_s=D c. Restore original
primal x=D u, equality dual y=R y_s and cone dual z=C z_s, also accounting for
any objective normalization. Scaling all coordinates of a cone by the same
positive factor preserves its cone. Row-specific SOC scaling is invalid.

This must be implemented as a separate solver representation. The current
RayGate raw dual audit retains original equality/cost data but treats live XC
as the original cone coefficients and offsets. Scaling XC in place would
silently change its audit. An immutable original cone snapshot, positive
inverse maps and original cost/rows must exist before transformation. Original
primal and dual gates must read that immutable representation after restoration.
Prepared lower-Fs refresh must update both raw and solver representations and
recompute Fs-dependent cone scales. Cache identity must include transformation
version, maps, Fs-dependent data, mesh/q/policy/phase and input/solve epochs;
saved physical fields must retain original variable layout. Numeric factors
must be invalidated whenever these data change.

These transformations are possible in VBA-owned arrays using the current DLL
operations. A DLL change would become useful for large snapshot/transform
kernels or exposing bounded refinement options, but is not mathematically
required. Any ABI extension should add a versioned entry point with validated
dimensions, positive finite factors, ownership/epoch and bounded-work contracts,
retain the existing entry points, and reproduce original operator residuals.
The existing native build uses MSVC with /fp:strict, warning-as-error, the VBA
floating-point compatibility assembly, and the repository's recorded licenses.

No free-problem solver, numerical tolerance, regularization or native binary
has been changed by this source audit. Native accuracy achievement and merge
remain gated by independent own-mesh q8 fields, original-model audits and
completed root metadata.
