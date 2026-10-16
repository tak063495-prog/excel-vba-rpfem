# Water input audit contracts

The JSON matrix tests ordinary public `RPX_SeepCheckInputs` and selected small
public `RPX_SeepRunSilent` cases. Positive controls prevent a validator from
passing merely by rejecting everything. The baseline runner does not replace
any published Fix3 product module.

| Input | Intended contract used by the audit |
| --- | --- |
| Water unit weight | Explicit positive finite number; blank retains 9.81. |
| Coupling / suction | Blank is zero; numeric 0/1, numeric text 0/1, logical Boolean values and text TRUE/FALSE have explicit logical meaning. Other explicit values are errors. |
| Material / region / point references | Positive whole ID, present in the appropriate main definition. Fractional IDs must not round to another existing reference. |
| Permeability | Positive finite kx; blank ky means kx. Very small positive values and anisotropy remain valid. Duplicate material entries are ambiguous and must be addressed explicitly. |
| Boundary head / water level / flux | Finite numeric value. Negative and zero values remain valid: an elevation datum and an outflow can be negative. |
| Boundary path | Existing endpoints on the selected boundary; start and end cannot accidentally mean the whole perimeter. The displayed circulation field is 0/1. Like main face paths, a traversal that crosses the end of the region loop requires circulation=1; a traversal that does not cross requires0. Blank circulation retains0. |
| Inverse unknown | `?` or `?` followed by a numeric initial value, including the supported full-width question mark. Arbitrary suffix text must not silently choose a different initial value. |
| Water table | At least two complete numeric x/y pairs, increasing distinct x. |
| Observations | Complete finite numeric pairs with x in the model range. Duplicate observed x values are not inherently invalid. |
| Table capacity | Rows 7 through 206. Data immediately beyond the documented table must not silently disappear. |
| Target | Zero delegates to main TARGET; otherwise integer 16 through 50000, matching the existing mesh generator. |
| krmin | Finite value strictly between zero and one. |
| Beta | Finite nonnegative value; zero remains a supported saturated linear control. |
| Relaxation | Finite value greater than zero and at most one. |
| Head tolerance | Positive finite value. |
| Iteration counts | Positive whole representable counts, not rounded fractions or text/error fallback. Mode3 inverse maximum must be at least3*(unknowns+1); one final forward solve is reserved and counted within the entered maximum. |
| Rigid permeability ratio | Positive finite value when used. No arbitrary upper limit or small-positive cutoff is added. |

Blank optional settings preserve their documented defaults. An explicitly
malformed value is distinct from a missing optional value. Numeric text remains
accepted where the existing numeric reader already permits it. Logical Booleans
are appropriate only to switches, not material IDs, permeability, gamma-w or
iteration counts.

These contracts do not change the flow equations, convergence threshold, GAP,
or valid negative head/flux signs. A rejected input is reported as an input
error rather than solver nonconvergence or a substitute physical model.
