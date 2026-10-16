# G1R20に対応するPython検証コード

結果だけでなく、実際に使った参照解法・改善手順・独立監査・失敗試行のコードをGitHubへ収録しています。
Python3.13、依存ライブラリは `validation/accuracy_20261013/requirements.txt` を使用します。
実験時はPython3.13.14、NumPy2.3.5、SciPy1.17.0、Clarabel0.11.1、各BLAS1threadでした。

| 用途 | GitHub内のコード |
|---|---|
| 26例の初期状態からの検証 | [run_26.py](../../validation/G1R19_python26_20261016/run_26.py)、[run_case.py](../../validation/G1R19_python26_20261016/run_case.py) |
| 元のq8組立・物理監査・根探索 | [python/](../../validation/G1R19_python26_20261016/python/) |
| P06の既知形状から新規求解 | [geometry_recipe_candidate.py](../../validation/G1R19_python26_geometry_replay_20261016/geometry_recipe_candidate.py) |
| 上限の原LP→capQP→同Fs再中心化 | [upper_goal_recipe.py](../../validation/G1R19_python26_upper_goal_recipe_20261016/upper_goal_recipe.py) |
| 元dualの1回線形補正 | [repair_original_dual.py](../../validation/G1R19_python26_dual_linear_repair_20261016/repair_original_dual.py) |
| P06A/Nの実行条件・形状由来・完了根 | [G1R20_RECIPE.md](../../validation/G1R19_python26_dual_linear_repair_20261016/G1R20_RECIPE.md) |
| 別プロセスの保存対再監査 | [audit_case.py](../../validation/G1R19_python26_20261016/audit_case.py) |
| ソース/入力/根/終了状態を分けた集計 | [collect_current.py](../../validation/G1R19_python26_20261016/collect_current.py)、[composite_provenance.py](../../validation/G1R19_python26_20261016/composite_provenance.py) |
| VBA出力とPythonの原問題・新規場の比較 | [audit_geometry_replay.py](tests/audit_geometry_replay.py) |
| 幾何だけの資産生成・実Excel比較 | [export_geometry_only.py](tests/export_geometry_only.py)、[run_geometry_replay_excel_v2.py](tests/run_geometry_replay_excel_v2.py) |
| ソース・単一DLL・配布物の照合 | [check_release_sources.py](tools/check_release_sources.py) |

初回の26例は次のように**未作成の**出力先へ実行します。

```powershell
python -X utf8 -m pip install -r validation/accuracy_20261013/requirements.txt
python -X utf8 validation/G1R19_python26_20261016/run_26.py --output validation/G1R19_python26_20261016/results/NEW_RUN
```

この元のcold driverだけでP06も自動達成するという意味ではありません。
実際に合格したP06は、上記の別版・形状再利用・新規求解・根の完了を含む手順です。
数値用の各source treeは試行時のbyteを保存し、後から元の実行コードを書き換えていません。
既知形状の元資産や採用・棄却した場は、別配布の証拠ZIPに含めています。

最適化をせず26対を再検査するには、
[EVIDENCE.zipと手順](../../validation/G1R19_python26_20261016/evidence/ALL_26_DELIVERY_20261016/README_JA.md)を使います。
これは保存場の再検査で、新規26計算の実行時間や正常終了を再測定する処理ではありません。

実Excelの再試験はWindows64bit Excelとpywin32を使用します。
専用Excelだけを起動する試験コードであり、検証用private bridgeは配布ブックへ組み込みません。
標準のVBA解析にPythonやVBAプロジェクト操作の許可は不要です。
