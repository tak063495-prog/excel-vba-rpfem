# rc9a verification (2026-10-21, Winmax2, local only)
- Build PASS (compile + reopen compile), workbook sha256 1bc0ae41b1cec8929fb9e166a360384786d79db172301ee6eccf86d34b0b00b1, DLL 66da1a21 unchanged.
- Repro (codex_fix9rc9a_repro/repro_rc9a.py, P02A_THIN short run, TARGET 320, thin upper 1 round, cap 800):
  - Scenario 0 (no injection): PASS, run_end PARTIAL_CERTIFIED, Fs 1.30309 / 1.31255.
  - Scenario A (thin upper trial: solver bypassed, iterate erased, HSD_SOLVE_BUDGET_EXCEEDED, iteration 0 = incident state): no error 9; fs_numerical_unknown reason=thin_budget_no_iterate endpoints_unchanged; round rejected reason=evaluate_failed (not fatal); run_end PARTIAL_CERTIFIED, Fs 1.30309 / 1.31576 written. (Driver flagged only its strict check original_trial_error: error logged at stage fix8_thin_evaluate rather than fs_trial.)
  - Scenario B (error 9 injected inside thin upper round): thin_layer_rejected reason=fatal -> adapt_fatal_preserved (lower 1.30309 / upper 1.31687, 324/324 el) -> run_end PARTIAL_CERTIFIED, Fs written, no later solves. PASS.
- Not run: full P06N/P06A, 7-case set (need user approval).
