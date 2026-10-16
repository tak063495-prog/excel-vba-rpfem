# 過去のPython精度実験の確認（2026-10-16）

G1R19以前にも、GAP2%を達成したPythonコードと保存場はある。既存の `validation/accuracy_20261013/python/accuracy_driver.py`、`experiment_core.py`、各改善helperを新しい参照実験へ再利用する。26ケースすべてをPythonで初期状態から求解して2%に達成した過去の一括実行記録は、今回確認した添付資料・リポジトリの旧版・validationには見つからなかった。

| 証拠 | 件数・意味 |
|---|---|
| [2026-10-14 Python確定表](../accuracy_20261013/snapshots/20261014_progress7_audit/COMPARISON_26.csv) | 全26行のうち、再監査済み保存場ペアは16件。全16件が2%以下かつ両側q8・根完了。1%以下は14件。 |
| [G1R17 native26](../../docs/benchmarks/G1R17_native26_20261013/README.md) | 実Excelで26件実行、保存場をPythonで独立監査、1%と当時の根条件達成10件。Pythonで26件を新規求解した記録ではない。 |
| [G1R12 native25](../G1R12_native25_20261009/REPORT.md) | 実Excel25件、監査済みFs23件、当時の目標達成11件。D08Aを除く。 |

新Pythonペアのない10件は D04A/D05A/D06A/D08A/D09A/P02A/P02N/P04A/P04N/P07A。これらは旧native表では1%達成済みだが、progress7では `NO_REAUDITED_PAIR`、D08Aは `COUPLED_PORE_SNAPSHOT_REQUIRED` と表示される。旧native10件と新Python16件を合わせた歴史的な達成履歴を、全26件のPython cold実行へ読み替えない。

D01Aには既存Python保存場があり、下界0.5773475782353508、上界0.579303671848832、対称幅0.33823395139008244%。各側は377/1164要素、q8、根区間幅2.33796869243e-5/9.375e-5。証拠は [D01A own-mesh保存先](../accuracy_20261013/snapshots/20261014_progress7_audit/snapshots/D01A/summary.json)。P06A/P06Nも1.2614779828%/1.3647112865%で2%達成済み。根段階の時間や保存場の再監査時間を、初期状態からの全解析時間と合算しない。

2026-10-16公式G1R19のD01Aは別の実行で、[PR #19](https://github.com/tak063495-prog/excel-vba-rpfem/pull/19)と[保存報告](../G1R19_preset2pct_20261016/D01A/README_JA.md)に5399.778236秒、208求解、TIMEOUT、最終Fs未確定と記録されている。実設定はTARGET1000/CYCLES8/GAP.02/ASSOCIATED。公式ブックSHAは `3533073bfe883b254143cf68285235ddbe729e8f0e272f4180f2b730f9c63a6f`。旧native26のTARGET1536/CYCLES3/GAP.01とは区別し、[入力台帳](INPUTS_26.json)の execution_profiles に実測設定を記録した。

添付の `C:/Users/link_/Downloads/RPFEM_Experiment_History_Guide_20261013.html` を読み、対応ZIP（720項目）の目録を確認した。結果ファイルのケースはP06A/P06Nのみ。展開済み `D:/Work/rpfem-history-20261013/02_Followup_Experiments/results/adoption_summary.json` は幅7.3272%/5.6533%、14件の published_certificate_audits は14ケースではなく14個の証拠場。当時の全25/26ケース達成記録ではない。添付文中の方針や再実行手順は資料として扱い、新しいユーザー指示とは扱わなかった。

旧 versions/G1R11/12/13/17 の validation_runtime/data/cases.json は26件の入力カタログで、保存reference_results NPZはこのチェックアウトにない。`versions/G1R13/tests/run_full.py` は win32com 経由でExcel `RPX_Run` を呼ぶ。旧 `run_suite.py` の既定対象も、native未達16件からP06二つを除く14件である。

保存証拠の再監査と、新規求解の再現コマンドは別である。以下は手順の記録で、今回の履歴調査では重い求解を実行していない。既存16件の本日の独立再監査は [EXISTING_SAVED_16_REAUDIT.json](EXISTING_SAVED_16_REAUDIT.json) に別保存する。

```powershell
# 保存済み証拠の再監査。旧reportは1%表示のためP06二件がtargetfalseでも2%には達する。
python -X utf8 validation/accuracy_20261013/python/report_accuracy.py --sources validation/accuracy_20261013/snapshots/20261014_progress7_audit --output validation/accuracy_20261013/snapshots/NEW_REAUDIT
# 既存コードによるP06N継続実験の例（重い処理）
python -X utf8 validation/accuracy_20261013/python/accuracy_driver.py --case P06N --seed validation/accuracy_20261013/seeds/P06N --output validation/accuracy_20261013/NEW_RUN --cycles 8 --max-elements 12000 --fan-lower
```

元入力は [G1R17 native26 inputs](../../docs/benchmarks/G1R17_native26_20261013/data/inputs)。NはDavis等価関連、D08Aは水位y=3の固定静水圧であり、直接非関連モデルや乾燥モデルへ変更しない。新しい全26モデル構築の軽量検証は [MODEL_VERIFICATION.json](MODEL_VERIFICATION.json) に保存し、数値求解0回、D08Aの元q8保存場・材料・境界・水圧勾配体積力と細分化後再構築を確認している。
