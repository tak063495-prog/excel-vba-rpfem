# G1R20 Fix9 rc9b 検証：P06A（Winmax2、2026-10-22、日本時間）

rc9b（ソース `51bde3d`（未公開のローカルコミット）、ブック SHA256 `f025cb2acac0e261f2333728677f96cd612d2de4adac30f2eac9f440a65fd788`、DLL `66da1a21…be70b` は Fix9 と同じ）の P06A の結果です。2% 標準プリセット、追加の設定は `ADAPT_THIN_UPPER_ROUND_MAX_ELEMENTS=8000`、`ADAPT_THIN_UPPER_CAP=9000`、`ADAPT_THIN_EARLY_UPPER_CAP=9000`、`ACCURACY_DLL_UPPER_EQUILIBRATE=1`（`inputs/P06A.json`）、制限時間 18171 秒。電源設定は変えていません。

| 状態 | Fs（下界〜上界） | GAP | 時間 |
|---|---|---:|---:|
| PASS_TARGET | 0.84307764（1481 要素）〜0.85815943（2944 要素） | 1.773% | 7470 秒 |

- 2026-10-22 00:10 開始、02:15 終了。エラーなし、独立監査 PASS、ソース・DLL のハッシュは実行前後で同じです。
- rc8 は 1.770%（7151 秒）でした。下界は同じで、上界がわずかに大きく（+3.0e-5）なっています。

ログは `cases/P06A/analysis.tsv`、結果は `results/P06A.json` です。

含めていないもの：`workbooks/`（実行後のブック 3.9 MB、DLL の複製、`RPFEM_logs/` の analysis.tsv は `cases/P06A/analysis.tsv` と同じ内容）、`prepared_template.xlsm`（2.4 MB）、`scripts/__pycache__/`。
