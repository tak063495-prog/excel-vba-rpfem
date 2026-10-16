# 26ケースの新規求解場・両根の独立監査

26/26ケースで原q8の物理上下限、各自己メッシュの根、対称GAP2%以内を確認しました。
正常な初期状態からのdriver終了まで確認できたのは20/26です。

| 区分 | ケース数 | 扱い |
|---|---:|---|
| 新規cold driver正常終了・独立監査合格 | 20 | cold終了確認済み |
| 新規再試行の場・両根合格、worker終了コード未記録 | 3 | D01A/D05A/D08A。正常終了時間を捏造しない |
| driver時間切れ、保存対・両根は独立監査合格 | 1 | P04N。正常cold終了には数えない |
| 既知メッシュ形状から新しい場・根を求解 | 2 | P06A/N。旧場・旧Fsは使わず、汎用cold生成と区別 |

P06Aは1.8173754281%、P06Nは1.9723805913%です。最後のP06N原dualはフルベクトルから
元問題を別プロセスで組み直して再監査しました。短い線形補正componentの事前PIDとOS過程時刻は
未記録なのでnullのまま保持し、実測wrapper時間は既知process時間合計と別項目です。

証拠は [G1R20 ReleaseのEVIDENCE.zip](https://github.com/tak063495-prog/excel-vba-rpfem/releases/download/G1R20/EVIDENCE.zip)。
98,287,686 byte、SHA256 `9121db1fa2f60ebdddf35d8c6123d4562cf631a3808c61cb02d2a36702e2960a`。
元1105ファイルとmanifest/REPLAYを収録し、元byteと絶対パス表記は資料として保持しています。
再生では相対フォルダーを使います。これは保存場の検査であり、新しい最適化やcold実行の再計測ではありません。

実験環境はPython3.13.14、NumPy2.3.5、SciPy1.17.0、Clarabel0.11.1、BLAS各1threadです。
ZIPを新しいフォルダーへ展開し、次のように未作成の出力先を指定します。

```powershell
python -X utf8 REPLAY.py --output replay_results
```

[集計CSV](../../CURRENT_RESULTS.csv)、[全試行・終了状態JSON](../../CURRENT_RESULTS.json)、
[再現手順](../../../G1R19_python26_dual_linear_repair_20261016/G1R20_RECIPE.md)を併せて参照してください。
独立展開先からの再生記録は `REPLAY_VERIFICATION.json` に保存します。
旧PARTIAL_24of26のZIP・記録はその時点の不変snapshotです。
