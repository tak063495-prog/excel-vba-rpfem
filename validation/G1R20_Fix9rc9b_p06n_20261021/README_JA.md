# G1R20 Fix9 rc9b 検証：P06N（Winmax2、2026-10-21〜22、日本時間）

rc9b（ソース `51bde3d`（未公開のローカルコミット）、ブック SHA256 `f025cb2acac0e261f2333728677f96cd612d2de4adac30f2eac9f440a65fd788`、DLL `66da1a21…be70b` は Fix9 と同じ）の P06N の結果です。2% 標準プリセット、追加の設定は `ADAPT_THIN_UPPER_ROUND_MAX_ELEMENTS=8000`、`ADAPT_THIN_UPPER_CAP=9000`、`ADAPT_THIN_EARLY_UPPER_CAP=9000`、`ACCURACY_DLL_UPPER_EQUILIBRATE=1`（`inputs/P06N.json`）、制限時間 18171 秒。電源設定は変えていません。

| 状態 | Fs（下界〜上界） | GAP | 時間 |
|---|---|---:|---:|
| FS_AVAILABLE_PARTIAL | 0.67486963（2000 要素）〜0.69145410（2938 要素） | 2.428% | 17690 秒 |

- 2026-10-21 19:15 開始、2026-10-22 00:10 終了。エラーなし、監査 PASS、ソース・DLL のハッシュは実行前後で同じです。
- Fs は rc8（2.428%、13552 秒）とビット一致です。上界の平衡化で上界の段 0〜7 の HSD への切り替えはなくなりましたが、Fs は変わりませんでした。4700 要素の上界の段は制限時間に達し、上界は更新されていません。下界は rc8 より遅くなりました（1480 要素の段 4258 秒、rc8 は約 2272 秒）。
- ログの `thin_layer_step seconds=-84046` は VBA の `Timer` が深夜 0 時に戻ったためと思われます（時間保護への影響は未確認）。

ログは `cases/P06N/analysis.tsv`、結果は `results/P06N.json` です。

含めていないもの：`workbooks/`（実行後のブック 4.1 MB、DLL の複製、`RPFEM_logs/` の analysis.tsv は `cases/P06N/analysis.tsv` と同じ内容）、`prepared_template.xlsm`（配布ブックから作る準備用、2.4 MB）、`scripts/__pycache__/`。
