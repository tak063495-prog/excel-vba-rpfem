`lower_polish_actual1481.npz` contains two actual unpolished Clarabel fields
saved during P06A experiments on 2026-10-13. Its embedded provenance identifies
the source files and their SHA-256 hashes. The mesh and physical model identities,
case, safety factors and q=8 accompany the original field arrays.

At Fs=0.834 the saved field fails the unchanged raw equilibrium tolerance
(1.0290023e-7 > 1e-7). At Fs=0.84 the other field fails the unchanged raw yield
tolerance. Tests freshly assemble the original physical programs and verify
both successful correction and rejection with the unchanged-source fallback.
These fields are test inputs; neither is a published accuracy certificate.
