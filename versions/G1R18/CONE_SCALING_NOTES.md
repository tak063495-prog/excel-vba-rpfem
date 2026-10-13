# Cone normalization integration investigation

This is a source investigation, not an implementation or numerical acceptance.
Normalize an entire scalar or SOC block by one finite strictly positive factor
`a = 1 / ||G_block||_F`, including its offset. A SOC's three rows must share
the same factor. Such normalization preserves the feasible primal set.

## Existing program and physical ownership

`RPX_Conic.RPX_Cone` and `RPX_LinearDirect` populate `XC` with coefficients and
offsets. `RPX_Compile` currently normalizes objective and equality rows, but
does not normalize cone blocks. `RPX_ConeData` has no normalization metadata.
`RPX_OriginalPrimalGate` and `RPX_ExportAssembled` operate on the original
assembled program before Compile. Keep that ordering so raw thresholds and
exported original programs remain authoritative.

`RPX_Affine` already preserves `originalCones=XC` before its equivalent Ruiz
cone/column scaling for fixed-load zero-objective witnesses. It restores the
primal and uses original cones in `RPX_AffineOriginalGate`. Adding ordinary
cone scaling after this original capture keeps that independent gate intact.

The physical lower/upper gates use the restored primal stress/velocity fields,
original material strengths, loads, geometry, and requested quadrature rather
than XC's transformed cone coefficients. Adaptive and immutable witness
payloads store those physical primal fields and their own mesh/context; they
need no rescaling or change in ownership.

## Compile, refresh and native boundaries

The initial safe normalization point is in `RPX_Compile`, before
`NativeCompilePattern`. Its `RPX_NativeConeIdentityArgs` calls `PrepareCones`,
which caches coefficients and offsets even though the compiler only requests
the identity arrays. Scaling afterward without `RPX_NativeBulkReset` would
leave the DLL using stale unscaled data.

Prepared Fs refresh needs explicit handling:

- `RPX_RefreshLower` overwrites offset/head coefficients with raw strength
  data but leaves SOC tail coefficients intact.
- `RPX_RefreshUpper` modelKind1 multiplies existing head coefficients by the
  strength ratio; modelKind2 rebuilds scalar rows with raw geometry/strength.
- Consequently, Compile-only normalization followed by either unchanged
  refresh mixes raw and scaled rows and changes the mathematical model.

For a reusable prepared program, preserve raw coefficients/offsets separately,
perform the existing strength refresh in raw units, recompute a positive block
factor, then derive every working coefficient/offset using that factor. Merely
undoing and reapplying previous scales is mathematically valid but accumulates
rounding, so deterministic raw regeneration is preferable for repeated roots.
Reset packed native coefficients after every such update.

`RPX_LowerCache.CopyCone` must copy new raw/factor metadata.
`RPX_LowerExactIdentity` must include the scaling policy so a compiled structure
cannot be mistaken for a different policy. Its current identity intentionally
excludes strength and quadrature; any factor dependence on strength must be
recomputed at refresh, including after restoring a cache built at a different
Fs. Ordering and symbolic patterns depend on variable identities, which do not
change under positive block scaling.

A conservative first implementation can bypass prepared refresh and the lower
model cache only while cone scaling is enabled, assembling each Fs from the
raw physical program. This trades speed for simpler correctness and leaves the
existing OFF path unchanged. Later reuse requires the raw/factor lifecycle
above and direct fresh-versus-prepared oracle comparisons.

## Dual and diagnostic units

For a transformed cone `G'=aG`, `h'=ah`, keep solver slack and dual in the
transformed units: `s'=a s`, `z'=z/a`. The original dual is `z=a z'`, and the
original slack is `s=s'/a`. The primal `x` is unchanged. Internal NT, HSD and
KKT operations must continue pairing the transformed coefficients with the
transformed dual; converting only the dual internally would break stationarity.

For raw dual certificates and external diagnostic comparisons, export the
factor per cone and map dual/slack explicitly. HSD failure capture currently
exports the compiled coefficients, candidate dual and slack; add factor and
policy metadata to make its units unambiguous. Scaling should not relax raw
physical tolerances or turn a failed optimum into a root certificate.

The existing native kernels accept flat numerical arrays and do not depend on
the VBA UDT layout. No DLL source or ABI change is necessary if normalization
and metadata lifecycle are handled in VBA before packing those arrays.

Required validation includes policy OFF invariance, positive factor/zero-row
guards, fresh versus prepared lower/upper at several Fs, cache on/off across
Fs, scalar and SOC3 blocks, pure VBA versus DLL, original raw primal/dual gates,
and preservation of input, mesh, and witness epochs.
