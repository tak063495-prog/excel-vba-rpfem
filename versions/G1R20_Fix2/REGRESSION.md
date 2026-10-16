# G1R20 Fix2 再現・回帰の実行方法

コマンドはGitHubリポジトリのルートで実行します。過去版・入力・Python参照を併用するため、
配布ZIPだけでは全検証を再実行できません。数値参照は
`validation/accuracy_20261013/requirements.txt`、Excelハーネスは加えて
`pywin32`、`psutil`、保存ソースの抽出は `oletools` を必要とします。
Windows・64bit Excelの専用コピーで実行し、Excelテストを同時実行しないでください。
出力フォルダーや結果JSONは新しい名前を指定します。

## ソース・監査の回帰

```powershell
python -X utf8 versions/G1R20_Fix2/tests/run_portable_regression.py --output <新しい結果.json>
python -X utf8 versions/G1R20_Fix2/tools/check_release_sources.py --require-delivery
python -X utf8 versions/G1R20_Fix2/tools/audit_book_sources.py versions/G1R20_Fix2
```

新しい状態管理の契約とFix1の97件を実行します。新規の数値解析回数ではありません。
63ソースのUTF8/CP932一致、Fix1との差分3モジュール、保存ブック内の全ソース、
単一DLLと配布ZIPを照合します。

## 実Excelでの状態管理

```powershell
python -X utf8 versions/G1R20_Fix2/tests/run_ui_setup_control_excel.py --execute --published-only --output <新しい負例フォルダー>
python -X utf8 versions/G1R20_Fix2/tests/run_ui_setup_control_excel.py --execute --output <新しい正例フォルダー>
python -X utf8 versions/G1R20_Fix2/tests/run_water_restart_excel.py --execute
python -X utf8 versions/G1R20_Fix2/tests/run_water_journal_dirty_excel.py --execute
```

初期化・メッシュ生成・プリセットの二重例外と、水理の保存・プロセス再起動を検査します。
製品オリジナルは変更せず、専用ブックのコピーへ試験コードを追加します。

Fix1で確認したプリセットと水理10条件も、新しい63ソースで検査します。

```powershell
python -X utf8 versions/G1R20_Fix2/tests/UI_run_excel.py --execute --output <新しいUI結果フォルダー>
python -X utf8 versions/G1R20_Fix2/tests/run_water_excel.py
```

## 新規の数値回帰

```powershell
python -X utf8 versions/G1R20_Fix2/tests/run_cold_regression_excel.py --case D02A --precision-from-start --budget 600
python -X utf8 validation/G1R19_D02_route_audit_20261016/audit_cold_native.py --run <D02A結果フォルダー> --case D02A --lower 00004_lower --upper 00005_upper --output <新しい監査.json>
python -X utf8 versions/G1R20_Fix2/tests/run_water_native_fixed_excel.py --seconds 240 --solve-seconds 20
```

D02Aは入力から新規に解析し、過去の場・メッシュ・Fs・根を読み込みません。
D08Aは規定水位結果と107要素を新規生成し、q8の下側Fs=1・上側Fs=2を解き、
独立Pythonで原q8を再組立して場を再監査します。D08Aは固定点の回帰です。

Release別添 `EXECUTION_EVIDENCE.zip` に、生ログ、試験ブック、ソーススナップショット、
原CSRと場を収録します。全ファイルのSHA256を `MANIFEST.json` に記録します。
