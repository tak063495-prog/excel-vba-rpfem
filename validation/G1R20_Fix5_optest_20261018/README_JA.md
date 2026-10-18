# G1R20 Fix5 動作テスト（Winmax2、2026-10-18）

G1R20 Fix5（`versions/G1R20_Fix5/delivery/RPFEM_20261018_G1R20_Fix5.xlsm`、SHA256 `601985a2…dd66f`、DLL `66da1a21…be70b`）を、2%標準プリセット・DLL 使用で実行した記録です。ハーネスは `validation/G1R20_Fix4_bench26_20261017` と同じもので、対象ケースと上限時間だけを変えています（`scripts/prepare.py`）。

| ケース | 判定 | Fs 下界〜上界 | GAP | 時間 |
|---|---|---|---|---|
| D05A | PASS_TARGET | 1.30142〜1.32029 | 1.44% | 187.8 s |
| P02A | PASS_TARGET | 1.30129〜1.32149 | 1.54% | 226.6 s |
| P05A | PASS_TARGET | 1.46421〜1.49342 | 1.98% | 1166.6 s |
| P10A（2回目、AC電源） | FS_AVAILABLE_PARTIAL | 2.00606〜2.04794 | 2.07% | 1686.3 s |
| P10A（1回目、バッテリー駆動） | FAILED `HSD_SOLVE_BUDGET_EXCEEDED` | – | – | 403.3 s |

- 上限時間：D05A・P02A 3600秒、P05A 7200秒、P10A 10800秒（Winmax2 用に余裕を持たせた値）。
- P10A の1回目は、最初の求解で HSD が300秒の上限に当たりました（バッテリー駆動中、約9秒/反復）。ユーザーの承認を得て AC 電源で再実行し、記録は `_run1_on_battery` の名前で残しています（`results/P10A_rerun_note.json`）。
- Fix4 の同じ Winmax2 での P10A 対照は `validation/G1R20_Fix4_control_P10A_20261018`（2.00606〜2.04794、2080.4秒）です。
- P06N/P06A はユーザーが別の PC で検証するため、実行していません。
- `workbooks/`（ケースごとのブックと生ログ）は容量のため載せていません。各ケースの `analysis.tsv`・`status.json`・witness は `cases/` にあります。
