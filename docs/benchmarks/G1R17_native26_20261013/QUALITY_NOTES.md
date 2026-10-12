# 採用精度と最終確認の記録

実行済みの診断ログを確認した記録。Fs取得と、指定した積分次数8を含む最終確認の完了を区別する。
上下界ギャップが1%以下でも、最終確認が未達なら目標達成とは扱わない。

| ケース | 判定 | 採用積分次数 下/上 | 精度・最終確認の記録 |
|---|---|---|---|
| D08A | 目標達成 | 8/8 | 指定目標と独立監査を達成 |
| P01A | Fs取得・精度目標未達 | 8/8 | GAP_TARGET_NOT_MET |
| P01N | Fs取得・精度目標未達 | 8/8 | GAP_TARGET_NOT_MET |
| P02A | 目標達成 | 8/8 | 指定目標と独立監査を達成 |
| P02N | 目標達成 | 8/8 | 指定目標と独立監査を達成 |
| P03A | Fs取得・精度目標未達 | 8/8 | GAP_TARGET_NOT_MET |
| P03N | Fs取得・精度目標未達 | 8/8 | GAP_TARGET_NOT_MET |
| P04A | 目標達成 | 8/8 | 指定目標と独立監査を達成 |
| P04N | 目標達成 | 8/8 | 指定目標と独立監査を達成 |
| P05A | Fs取得・精度目標未達 | 8/8 | GAP_TARGET_NOT_MET |
| P05N | Fs取得・精度目標未達 | 8/8 | GAP_TARGET_NOT_MET |
| P07A | 目標達成 | 8/8 | 指定目標と独立監査を達成 |
| P07N | Fs取得・精度目標未達 | 8/8 | GAP_TARGET_NOT_MET |
| P08A | Fs取得・精度目標未達 | 4/4 | FINAL_NUMERICAL_UNKNOWN;OLD_WITNESSES_KEPT HSD_SOLVE_BUDGET_EXCEEDED |
| P09A | Fs取得・精度目標未達 | 4/4 | FINAL_NUMERICAL_UNKNOWN;OLD_WITNESSES_KEPT HSD_SOLVE_BUDGET_EXCEEDED |
| P10A | Fs取得・精度目標未達 | 4/4 | FINAL_NUMERICAL_UNKNOWN;OLD_WITNESSES_KEPT HSD_SOLVE_BUDGET_EXCEEDED |
| D01A | Fs取得・精度目標未達 | 8/8 | GAP_TARGET_NOT_MET / LOWER_ROOT_NOT_VALID / UPPER_ROOT_NOT_VALID |
| D02A | Fs取得・精度目標未達 | 4/8 | GAP_TARGET_NOT_MET / LOWER_ROOT_NOT_VALID / UPPER_ROOT_NOT_VALID |
| D03N | Fs取得・精度目標未達 | 8/8 | GAP_TARGET_NOT_MET / LOWER_ROOT_NOT_VALID / UPPER_ROOT_NOT_VALID |
| D04A | 目標達成 | 8/8 | 指定目標と独立監査を達成 |
| D05A | 目標達成 | 8/8 | 指定目標と独立監査を達成 |
| D06A | 目標達成 | 8/8 | 指定目標と独立監査を達成 |
| D07A | Fs取得・精度目標未達 | 8/4 | NATIVE_GLOBAL_INCOMPLETE |
| D09A | 目標達成 | 8/8 | 指定目標と独立監査を達成 |
| P06N | Fs取得・精度目標未達 | 8/8 | GAP_TARGET_NOT_MET |
| P06A | Fs取得・精度目標未達 | 8/8 | GAP_TARGET_NOT_MET |
