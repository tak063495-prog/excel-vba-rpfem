# 検証の再現手順

## Pythonの独立検証

Python 3.12の作業環境へ`requirements.txt`の依存をインストールします。検証資料を変更しないよう、一式の作業用コピーで実行してください。通常のExcel利用にPythonは不要です。

```text
python -m pip install -r requirements.txt
python python/diagnose.py --tag new_diagnosis
python python/paired_python.py
python python/fan_repair.py
python python/compare_assembled.py
python python/verify_coarse_python.py
python python/test_meshfan_native.py
```

`diagnose.py`は保存済み実メッシュ243要素を読み、固定Fsの下界・上界と外接多角形LPを解きます。出力先の既存フォルダーを上書きしません。別の`--tag`を指定します。`paired_python.py`はP06A/P06Nの元メッシュ・修復メッシュを同じ材料・荷重で照合します。`compare_assembled.py`は実Excelが出力したSOCP係数とPython組立を比較します。`verify_coarse_python.py`は保存済み実Excel結果の全応力・速度場とFs根区間を独立監査します。

Pythonのモデル・監査コードは`validation_runtime/`、元メッシュの読み取り記録は`data/P06N_original243_cells.json`、ケース入力は`data/inputs/`に同梱しています。DLLの直接検証はWindows x64で実行します。

Clarabelの`AlmostSolved`や`NumericalError`を無条件で合格扱いにしていません。物理監査、双対制約、求解の役割を別々に記録します。P06A下界根の反対側では、最初の不可行証明の絶対残差が不足したため、別途、最大荷重の双対側を検査しました。その点のPython主変数場は監査不合格のまま記録し、採用しません。

## 実Excelの回帰

再実行にはWindows x64 Excelと、このリポジトリのG1R16 DLL5基準ファイルが必要です。自動導入・試験モジュールの追加には、利用者が既に認めているVBAプロジェクトアクセス環境が必要です。ツールは許可・セキュリティを自動変更しません。Excel試験を並行起動しないでください。

```text
python tools/run_coarse.py --case P06N
python tools/run_coarse.py --case P06A
python tools/acceptance.py
python tools/test_original_P06A.py
python tools/reopen_delivery.py
python tools/verify_delivery.py
```

試験ツールは既存出力がある場合に拒否するものがあります。保存済み証拠を退避した別の作業用コピーで再実行します。`acceptance.py --resume`はDLLのハッシュが一致する場合だけ完了試験を引き継ぎます。試験コードだけを修正した履歴と、失敗した試験ログは保持します。

`test_geometry.py`は25ケースの**メッシュ生成だけ**を比較します。25ケースのFs全解析ではありません。`test_fixed_precision.py`はP06N Fs=0.5、P06A Fs=0.7の固定点のみです。`run_coarse.py`はTARGET200、q8、適応OFFで、フル解析を開始しません。

`test_guards.py`の注入コードは隔離した試験ブックだけに置きます。製品ブックに試験フックを含めません。実際のC++メモリ不足と未知例外は`tests/build_meshfan_failure_test.cmd`で確認できます。

DLLの再ビルドは`native/build.cmd`を使用します。MSVC x64のパスは実行環境に合わせて確認してください。`/fp:strict`と既存のx87式・Double保存境界を維持します。完成したDLLを再ビルドするとハッシュが変わるため、そのバイナリに対して回帰を再実行します。
