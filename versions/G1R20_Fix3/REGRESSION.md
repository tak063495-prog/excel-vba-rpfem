# G1R20 Fix3 再現・回帰の実行方法

リポジトリのルートで実行します。過去版・入力・Python数値参照も使用するため、全検証にはGit cloneが必要です。
数値参照は `validation/accuracy_20261013/requirements.txt`、Excel試験はWindowsの64bit Excel・pywin32・psutil、
保存ソースの抽出はoletoolsを必要とします。Excel試験は同時実行せず、専用コピーで実施してください。
出力パスには既存ファイル／フォルダーを指定しません。

```powershell
python -X utf8 versions/G1R20_Fix3/tests/run_portable_regression.py --output <新規結果.json>
python -X utf8 versions/G1R20_Fix3/tools/check_release_sources.py --require-delivery
python -X utf8 versions/G1R20_Fix3/tools/audit_book_sources.py versions/G1R20_Fix3
```

全63標準モジュールとThisWorkbookイベントのUTF8/CP932・保存ブック・DLL・配布ZIPを照合します。
portableはソース／拒否契約であり、新規数値求解の件数ではありません。

```powershell
python -X utf8 versions/G1R20_Fix3/tests/run_saved_formula_excel.py --execute --negative --output <新規負例>
python -X utf8 versions/G1R20_Fix3/tests/run_saved_formula_excel.py --execute --output <新規正例>
python -X utf8 versions/G1R20_Fix3/tests/run_saved_formula_excel.py --execute --scenario missingproof --output <新規未確認proof結果>
python -X utf8 versions/G1R20_Fix3/tests/run_saved_formula_excel.py --execute --scenario openchanged --output <新規再開時数式変化結果>
python -X utf8 versions/G1R20_Fix3/tests/run_read_model_fix3.py --phase before
python -X utf8 versions/G1R20_Fix3/tests/run_read_model_fix3.py --phase fixed
python -X utf8 versions/G1R20_Fix3/tests/UI_run_excel.py --execute --output <新規UI結果>
python -X utf8 versions/G1R20_Fix3/tests/run_ui_setup_control_excel.py --execute --output <新規制御結果>
```

```powershell
python -X utf8 versions/G1R20_Fix3/tests/run_seep_ui_fix3_excel.py --execute --negative --output <新規浸透UI負例>
python -X utf8 versions/G1R20_Fix3/tests/run_seep_ui_fix3_excel.py --execute --output <新規浸透UI正例>
python -X utf8 versions/G1R20_Fix3/tests/run_water_source_fix3_excel.py --execute --negative --output <新規水理負例>
python -X utf8 versions/G1R20_Fix3/tests/run_water_source_fix3_excel.py --execute --output <新規水理正例>
python -X utf8 versions/G1R20_Fix3/tests/run_water_source_process_restart.py --execute --output <新規水理プロセス再開>
```

savedformulaの追加シナリオは `--scenario missingproof`、`--scenario openchanged` で指定します。
浸透UI・全modeのsource／保存再起動テストの詳細はtests内のprobeと証拠記録を参照してください。
保存表示テストは合成sentinelを使い、数値証明を生成したとは扱いません。

```powershell
python -X utf8 versions/G1R20_Fix3/tests/run_cold_regression_excel.py --case D02A --precision-from-start --budget 600
python -X utf8 validation/G1R19_D02_route_audit_20261016/audit_cold_native.py --run <新規D02A結果> --case D02A --lower 00004_lower --upper 00005_upper --output <新規監査.json>
python -X utf8 versions/G1R20_Fix3/tests/run_water_native_fixed_excel.py --seconds 240 --solve-seconds 20
```

D02Aは入力からの新規解析で、過去の場・メッシュ・Fs・根を読みません。
D08Aは規定水位と107要素を新規生成し、q8の下側Fs=1・上側Fs=2を解く固定点の回帰です。
Pythonで原q8を独立再組立してVBA/DLLの場を監査します。D08Aの根探索や本番全ケース完走とは区別します。
Release別添 `EXECUTION_EVIDENCE.zip` に、試験用コピー・ログ・場・ソーススナップショットとSHA256を収録します。
