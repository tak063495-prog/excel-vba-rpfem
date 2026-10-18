# G1R20 Fix7 rc3 検証（不採用、Winmax2、2026-10-18）

rc3 は、rc2 の `RPX_HsdOrdinaryLimit` の倍率を 1.5 から `ADAPT_ACCURACY_ESTIMATE_FACTOR`（既定 1.2）に変えた候補です（SOURCE_COMMIT `66eaed8`、このリポジトリの公開履歴には含まれていません。差分は `rc3_RPX_Solve_vs_rc2.diff` を見てください）。

P10A が **FS_AVAILABLE_PARTIAL**（2.00606〜2.05088、上界 427 要素、GAP 2.210%、1113.1 秒）となり精度が落ちたため、**不採用** です。
582 要素のラウンドの見積もり 363.5 秒が上限 360 秒を超え、そこで細分化を止めてしまいました。
D05A 239.1 秒、P02A 295.4 秒、P05A 1469.7 秒（いずれも Fs は Fix6 と同じ）。
`winmax2_chain_state.txt` は実行時の進捗記録です。
