# G1R20 Fix4 26ケース速度計測（Triela, 2%・速度優先プリセット, 2026-10-17）

G1R20 Fix4（`versions/G1R20_Fix4`、commit `3a4f3e67850969fef5d4ae442149d07d7f720ba8`）の配布ブックに、2%・速度優先プリセット `RPX_Speed2pctPreset` を適用し、26ケースを実Excelで1ケースずつ順番に解析した記録です。DLL は `DIRECT_DLL=1` で使用しています。

- PC: Triela（AMD Ryzen 9 5900X 12コア/24スレッド、Windows 11 Pro、64bit Excel 16.0）
- 実行: 2026-10-17 11:03 JST 開始 → 2026-10-18 00:25 JST 全26ケース終了（約13.4時間）
- 配布ブック SHA256: `b010ab89c4e70540277475d1aab025db2545a24fd6697ad87856998737e94be6`（`versions\G1R20_Fix4\delivery\RPFEM_20261017_G1R20_Fix4.xlsm`）
- DLL SHA256: `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`（`rpfem_native_x64.dll`）
- 実行中に配布ブックと DLL の SHA が変わっていないことを確認しました（`results/ALL26_FINISHED.json` の `source_sha_unchanged` / `dll_sha_unchanged`）。
- 時間は `run_end.elapsed_s`（解析本体）です。Excel起動・保存・独立監査は含みません。

## 結果のまとめ

- 判定: PASS_TARGET 21、FS_AVAILABLE_PARTIAL 3（P10A・P06N・P06A）、TIMEOUT_STOPPED 1（P05A）、FAILED 1（D05A）。強制終了（KILLED）は0件です。
- **P05A** は 1200 秒の上限ちょうどで停止しました。停止直前の判定は PASS_TARGET で、GAP 1.98%、独立監査も合格です。ただし上限に達したため TIMEOUT_STOPPED としています。
- **P10A** は GAP 2.07% で、2% にわずかに届きませんでした（2262 秒）。
- **P06N** は GAP 30.5%（16193 秒）で、2% に大きく届きませんでした。詳細は [#29](https://github.com/tak063495-prog/excel-vba-rpfem/issues/29) にあります。P06A も GAP 24.7%（16317 秒）で同じ傾向です。
- **D05A**（座標1000倍・γ1000分の1の相似ケース）は 11 秒で `HSD_NOT_CONVERGED` により失敗しました。実際には HSD は OPTIMAL で、その後の dual 復元の不採用がこの名前で報告されています。詳細は [#28](https://github.com/tak063495-prog/excel-vba-rpfem/issues/28) にあります。
- 合計時間: Fix4 47,144 秒、G1R17 DLL5 27,967 秒（比 1.69）、G1R19 23,997 秒（比 1.96）。
- P06N・P06A の2ケースだけで Fix4 は 32,510 秒かかっています。この2ケースを除くと Fix4 14,634 秒、G1R17 15,860 秒（比 0.92）、G1R19 14,656 秒（比 1.00）です。G1R17・G1R19 は GAP 1% 目標の別プロファイルで、多くのケースが目標未達のまま終わっているため、時間だけの比較には注意が必要です。

## ケース別

比 = Fix4 の時間 ÷ 比較対象の時間（1 未満なら Fix4 が速い）。

- G1R17 DLL5: [docs/benchmarks/G1R17_native26_20261013](../../docs/benchmarks/G1R17_native26_20261013)（同じ Triela / 5900X、2026-10-12〜13、GAP 1% 目標の本番プロファイル）
- G1R19: 同じ DLL（SHA `66da1a21…`）による 2026-10-16 の native26 計測（G1R17 本番プロファイル、GAP 1% 目標）。プリセットが違うため参考値です。

| ケース | 判定 | Fs 下界 | Fs 上界 | GAP | 独立監査 | DLL_READY | 要素数（下/上） | 時間 s | 上限 s | G1R17 s | 比 | G1R19 s | 比 | 備考 |
|---|---|---:|---:|---:|---|---|---:|---:|---:|---:|---:|---:|---:|---|
| D08A | PASS_TARGET | 1.3016 | 1.3260 | 1.86% | 合格 | ○ | 429 / 429 | 726.4 | 900 | 348.1 | 2.09 | 220.2 | 3.30 |  |
| P01A | PASS_TARGET | 0.6707 | 0.6830 | 1.82% | 合格 | ○ | 432 / 432 | 482.7 | 900 | 264.6 | 1.82 | 203.7 | 2.37 |  |
| P01N | PASS_TARGET | 0.6649 | 0.6772 | 1.82% | 合格 | ○ | 428 / 428 | 458.2 | 1200 | 262.8 | 1.74 | 198.5 | 2.31 |  |
| P02A | PASS_TARGET | 1.3013 | 1.3215 | 1.54% | 合格 | ○ | 355 / 355 | 234.4 | 1200 | 334.1 | 0.70 | 227.6 | 1.03 |  |
| P02N | PASS_TARGET | 1.2149 | 1.2364 | 1.75% | 合格 | ○ | 355 / 355 | 243.5 | 1200 | 321.4 | 0.76 | 223.5 | 1.09 |  |
| P03A | PASS_TARGET | 1.1534 | 1.1751 | 1.87% | 合格 | ○ | 587 / 587 | 883.8 | 1200 | 252.9 | 3.49 | 166.6 | 5.31 |  |
| P03N | PASS_TARGET | 1.1500 | 1.1718 | 1.88% | 合格 | ○ | 587 / 587 | 850.1 | 1200 | 254.1 | 3.35 | 190.7 | 4.46 |  |
| P04A | PASS_TARGET | 2.2491 | 2.2848 | 1.57% | 合格 | ○ | 355 / 355 | 224.8 | 1200 | 336.2 | 0.67 | 221.8 | 1.01 |  |
| P04N | PASS_TARGET | 2.1374 | 2.1745 | 1.72% | 合格 | ○ | 355 / 355 | 239.5 | – | 310.3 | 0.77 | 253.5 | 0.94 | 完了直後に runner が一時停止（下記） |
| P05A | TIMEOUT_STOPPED | 1.4642 | 1.4934 | 1.98% | 合格 | ○ | 656 / 656 | 1,201.0 | 1200 | 283.9 | 4.23 | 601.6 | 2.00 | 1200 s 上限で協調停止（停止前の判定 PASS_TARGET、GAP 1.98%、監査合格） |
| P05N | PASS_TARGET | 1.4561 | 1.4855 | 1.99% | 合格 | ○ | 656 / 656 | 1,331.8 | 1659 | 552.9 | 2.41 | 377.4 | 3.53 |  |
| P07A | PASS_TARGET | 1.3745 | 1.3978 | 1.69% | 合格 | ○ | 367 / 367 | 284.3 | 1644 | 547.5 | 0.52 | 996.4 | 0.29 |  |
| P07N | PASS_TARGET | 1.3255 | 1.3491 | 1.76% | 合格 | ○ | 367 / 367 | 294.9 | 2241 | 746.9 | 0.39 | 1,017.9 | 0.29 |  |
| P08A | PASS_TARGET | 1.4445 | 1.4628 | 1.26% | 合格 | ○ | 4018 / 374 | 1,297.9 | 4125 | 1,375.3 | 0.94 | 964.1 | 1.35 |  |
| P09A | PASS_TARGET | 0.9655 | 0.9847 | 1.97% | 合格 | ○ | 4018 / 450 | 2,130.3 | 3381 | 1,127.4 | 1.89 | 980.6 | 2.17 |  |
| P10A | FS_AVAILABLE_PARTIAL | 2.0061 | 2.0479 | 2.07% | 合格 | ○ | 4018 / 582 | 2,261.7 | 2970 | 990.0 | 2.28 | 814.6 | 2.78 | GAP 2.07%（わずかに未達） |
| D01A | PASS_TARGET | 0.5773 | 0.5774 | 0.01% | 合格 | ○ | 369 / 8 | 254.1 | 5589 | 1,862.8 | 0.14 | 1,787.2 | 0.14 |  |
| D02A | PASS_TARGET | 1.5000 | 1.5001 | 0.00% | 合格 | ○ | 360 / 8 | 146.0 | 6171 | 2,057.0 | 0.07 | 2,012.8 | 0.07 |  |
| D03N | PASS_TARGET | 1.1180 | 1.1181 | 0.01% | 合格 | ○ | 360 / 8 | 146.6 | 6105 | 2,035.1 | 0.07 | 2,182.7 | 0.07 |  |
| D04A | PASS_TARGET | 1.3013 | 1.3215 | 1.54% | 合格 | ○ | 355 / 355 | 234.5 | 1200 | 322.3 | 0.73 | 206.9 | 1.13 |  |
| D05A | FAILED | – | – | – | – | ○ | – | 11.4 | 1200 | 363.8 | 0.03 | 239.7 | 0.05 | 11 s で失敗 HSD_NOT_CONVERGED（[#28](https://github.com/tak063495-prog/excel-vba-rpfem/issues/28)） |
| D06A | PASS_TARGET | 1.3012 | 1.3215 | 1.55% | 合格 | ○ | 355 / 355 | 241.4 | 1200 | 321.3 | 0.75 | 207.1 | 1.17 |  |
| D07A | PASS_TARGET | 1.2643 | 1.2809 | 1.31% | 合格 | ○ | 355 / 355 | 220.0 | 1200 | 265.8 | 0.83 | 153.2 | 1.44 |  |
| D09A | PASS_TARGET | 1.3013 | 1.3215 | 1.54% | 合格 | ○ | 355 / 355 | 234.7 | 1200 | 323.3 | 0.73 | 207.1 | 1.13 |  |
| P06N | FS_AVAILABLE_PARTIAL | 0.6212 | 0.8447 | 30.49% | 合格 | ○ | 1290 / 1290 | 16,193.0 | 18171 | 6,057.4 | 2.67 | 4,490.9 | 3.61 | GAP 30.5%（[#29](https://github.com/tak063495-prog/excel-vba-rpfem/issues/29)） |
| P06A | FS_AVAILABLE_PARTIAL | 0.7758 | 0.9950 | 24.75% | 合格 | ○ | 947 / 947 | 16,317.0 | 18150 | 6,049.8 | 2.70 | 4,850.7 | 3.36 | GAP 24.7%（#29 と同じ傾向） |

- GAP は対称相対幅（`relative_gap`）です。判定は `scripts/grade_results.py` によります（根拠: `RPX_Fs.bas:572-573,613`、`RPX_Guard.bas:372-381`、独立の物理場監査）。
- D01A・D02A・D03N の「要素数（上）」欄は記録値が 8 です。c0 直接上界の経路で、要素数ではない値が入っていると推定しています（未確認）。
- 上限時間は過去の時間の約3倍（最低 900〜1200 秒）です。途中で runner を改訂して上限を見直したため、ケースにより 900 秒と 1200 秒が混在しています（`scripts/previous_revisions/`）。

## 実行中の出来事

- 11:24 JST: 上限時間の見直しのため、P01N 実行中に runner を入れ替えて再起動しました（`results/handoff_log.txt`）。P01N は最初から解析し直しています。
- 12:14 JST: P04N の終了処理中に `cases\P04N\status.json` の読み取りで PermissionError が出て、runner が止まりました（`results/batch_stdout.txt`）。P04N 自体は PASS_TARGET で完了済みで、12:26 JST に次のケースから再開しました。このため `results/P04N.json` には `timed_out` など一部の欄がありません。
- 他の Excel プロセスとの同時実行はありません（`results/other_excel_at_batch_start_*.json`）。

## 有効な設定

`RPX_Speed2pctPreset` の設定値です（`case_manifest.json` の `preset_expected`。各ケースで全29項目の一致を確認）。

`MODE=Fs`, `TARGET=1000`, `QUADRATURE=8`, `FS_TOL=0.0001`, `GAP=0.02`, `ADAPT=1`, `CYCLES=8`, `RAW_LOWER_AUDIT=1`, `C0_AUTO=0`, `ADAPT_ACCURACY_CAP=24000`, `ADAPT_ACCURACY_ROUNDS=8`, `ADAPT_LOWER_TARGET=4000`, `ADAPT_LOWER_FAN_ROUNDS=1`, `ADAPT_CERT_QUADRATURE=8`, `ACCURACY_DLL_RECOVERY=1`, `ACCURACY_DLL_SECONDS=120`, `ACCURACY_DLL_DUAL_SECONDS=20`, `ACCURACY_DLL_CONE_SCALING=1`, `HSD_SOLVE_SECONDS=300`, `DIRECT_DLL=1`, `ADAPT_PRECISION_FROM_START=1`, `ADAPT_C0_WEDGE_AUTO=1`, `ADAPT_BALANCED_LOWER_TARGET=0`, `ADAPT_UPPER_FLOW_PASSES=0`, `FS_TERMINAL_SCHEDULE=0`, `ACCURACY_DLL_UPPER_RECENTER=0`, `ACCURACY_DLL_DUAL_LINEAR_REPAIR=0`, `ACCURACY_DLL_DUAL_REGULARIZATION=0`, `DIRECT_DLL_SECONDS=5`

ケースごとの入力値（重力・地震係数・解析方針など）は `inputs/*.json` から設定しています。各ケースの実際の設定全体は `results/<ケース>.json` の `actual_settings` にあります。

## ファイル

- `results/progress.csv`、`results/batch_summary.json`、`results/ALL26_FINISHED.json`、`results/environment.json`
- `results/<ケース>.json`: ケースごとの結果（設定、Fs、監査、DLL カウンタ、フェーズ別時間、各ケースのブック SHA）
- `results/batch_stdout*.txt`: runner の出力
- `cases/<ケース>/`: `status.json`・`stage.txt`・witness（D05A は `analysis.tsv` も）。他のケースの `analysis.tsv`（P06N 47 MB、P06A 36 MB など）は大きいため収録せず、Triela に保管しています。
- `scripts/`、`run_batch.cmd`、`handoff.ps1`、`tests/BatchEvidence.bas`: 計測に使ったスクリプト
- `inputs/`、`case_manifest.json`: 入力と実行計画
- `harness_validation/`: 計測前の試運転（設定欠落の検出で停止した記録）
- `SHA256SUMS.txt`: このフォルダ内の全ファイルの SHA256
- 解析用ブック（`.xlsm`、各約2.3〜3.7 MB）は収録していません。
