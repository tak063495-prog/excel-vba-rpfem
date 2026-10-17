# Public G1R20 EVIDENCE.zip verification

Status: PASS. Downloaded directly from the public release URL on 2026-10-17T11:38:44.063475+00:00.

- URL: https://github.com/tak063495-prog/excel-vba-rpfem/releases/download/G1R20/EVIDENCE.zip
- Bytes: 98,287,686; expected 98,287,686.
- SHA256: `9121db1fa2f60ebdddf35d8c6123d4562cf631a3808c61cb02d2a36702e2960a`; matches the published original record.
- ZIP entries: 1107; CRC check passes.
- Original manifest files: 1105; every SHA256 matches.
- P06N composite mesh/field/certificate/summary/component entries: all eight present, manifest hashes match; available tracked originals also match.
- P06N component source-map checks: 542 match; 0 referenced source entries not present in this ZIP.
- P06N original dual and related members: 46; their recorded manifest hashes match.

P06N historical Python composite has L=0.6767771508826014, U=0.6902587260388777, symmetric GAP=1.9723805913033%, 5966/8739 elements. It imports historical geometry and computes new fields/roots; it is not generic cold execution. This check verifies archive bytes and source/data provenance, not physical feasibility by itself. Optimizer calls: zero.

The new G1R20 Fix4 30.5% run is a different 2026-10-17 speed-preset run. Its status/analysis/stage raw files were not found in the scoped public/local search. Values such as 19 timeouts and 4370 seconds of own-q8 work must be attributed to the issue reporter's comment, not claimed as independently recomputed from this old evidence ZIP. Private or newly published logs may exist outside the search scope.

Full member hashes and provenance are in PUBLIC_ARCHIVE_VERIFICATION.json.

Additional reference closure: 8 original summary references match; all four composite NPZs are bytewise identical to their adopted original component files. The independent original-dual audit's 58 before/after/code references also match the ZIP. The 2,592,313-byte repaired original full-dual NPZ has SHA256 `416797367e8eb700067dbeb0d8e7c1d019ecfb78dec166e8c1751fdad152c34a` and is bytewise identical to the saved root guide. No scientific audit or optimization was run here.
