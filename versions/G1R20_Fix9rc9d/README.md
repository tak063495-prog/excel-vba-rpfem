# G1R20 Fix9 rc9d candidate (local, unpublished)

rc9d = rc9c (f7394af) + ACCURACY_DLL_UPPER_EQUILIBRATE default changed 1 -> 0
(code default in RPX_AccuracyNative, settings row in RPX_UI, and both RPX_UI presets).
4 ASCII lines changed; 731 non-ASCII lines identical; cp932 imports re-encoded.

Reason (2026-10-22 JST): P02A_THIN eq0 on rc9c gives PASS_TARGET with GAP 0.457% in 1100 s,
identical to the Fix9 reference path. With eq1 it misses at GAP 0.531% in 1484 s, diverging from adapt cycle 1.
Value 1 is still available via the settings row.

Build status: pending. A waiter builds it after the P06N lower-only measurement finishes (see TEST_RESULTS.md when done).
