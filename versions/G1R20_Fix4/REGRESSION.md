# G1R20 Fix4 再現・回帰の実行方法

リポジトリのルートで実行します。過去版のブック、入力、Python数値参照も使うため、全検証にはGit cloneが必要です。
数値参照は `validation/accuracy_20261013/requirements.txt`、Excel試験はWindowsの64bit Excel・pywin32・psutil、保存ソース抽出はoletoolsを必要とします。
Excel試験は同時実行せず、ユーザーのExcelを閉じた専用環境で行います。各runnerは既存Excelを検出すると開始しません。出力先には新規パスを指定してください。

```powershell
python -X utf8 versions/G1R20_Fix4/tests/run_portable_regression.py --output <新規結果.json>
python -X utf8 versions/G1R20_Fix4/tests/run_final_record_checks.py --output <新規ハーネス結果.json>
python -X utf8 versions/G1R20_Fix4/tests/run_evidence_packager_checks.py --output <新規梱包検査.json>
python -X utf8 versions/G1R20_Fix4/tools/check_release_sources.py --require-delivery
python -X utf8 versions/G1R20_Fix4/tools/audit_book_sources.py versions/G1R20_Fix4
```

portableは過去版の125契約とFix4の入力・幾何検査を実行します。ソース検査と独立の矩形・凹形状oracleを含み、新規数値求解の件数ではありません。
ハーネス検査は後段失敗時の24件と記録集約の12件を合成して確認します。Excel・求解は実行せず、132件のportableとは別枠です。ローカルrawを再監査する場合は `--audit-records` を追加します。
梱包検査は8件を別の一時ROOTで実行し、元FAILの保持、後日終了観測、PID/起動時刻一致、生存プロセスの拒否を確認します。実Excelの起動・終了操作は行いません。

```powershell
# 既定は公開済みFix3の負例。正例は最終保存ブックをそのまま使います。
python -X utf8 versions/G1R20_Fix4/tests/run_seismic_phase.py --input-matrix --legacy
python -X utf8 versions/G1R20_Fix4/tests/run_seismic_phase.py --final-book --expected-book-sha b010ab89c4e70540277475d1aab025db2545a24fd6697ad87856998737e94be6 --input-matrix --legacy
python -X utf8 versions/G1R20_Fix4/tests/run_seismic_phase.py --production
python -X utf8 versions/G1R20_Fix4/tests/run_seismic_phase.py --final-book --expected-book-sha b010ab89c4e70540277475d1aab025db2545a24fd6697ad87856998737e94be6 --production

python -X utf8 versions/G1R20_Fix4/tests/run_settings_input_fix4.py --execute --output <新規設定結果>
python -X utf8 versions/G1R20_Fix4/tests/run_seep_state_fix4.py --execute --output <新規失敗時状態結果>
python -X utf8 versions/G1R20_Fix4/tests/run_saved_revision_fix4.py --execute --task migration --output <新規旧版移行結果>
python -X utf8 versions/G1R20_Fix4/tests/run_saved_revision_fix4.py --execute --task pure_public --output <新規保存ブック公開API結果>
python -X utf8 versions/G1R20_Fix4/tests/run_settings_extent_display_fix4.py --execute --output <新規設定表外・保存表示結果>
```

土の保存表示失効には合成sentinelを使います。sentinelを数値証明として扱いません。地震の履歴回帰とpure_publicは実際のFsを求めます。

```powershell
python -X utf8 versions/G1R20_Fix4/tests/run_water_input_matrix_excel.py --execute --negative --output <新規水理入力負例>
python -X utf8 versions/G1R20_Fix4/tests/run_water_input_matrix_excel.py --execute --output <新規水理入力正例>
python -X utf8 versions/G1R20_Fix4/tests/run_water_fix4_excel.py --execute --output <新規中止・単位体積重量結果>
python -X utf8 versions/G1R20_Fix4/tests/run_water_transition_excel.py --execute --output <新規消費・移行結果>
python -X utf8 versions/G1R20_Fix4/tests/run_water_inverse_budget_excel.py --execute --output <新規逆解析回数結果>
python -X utf8 versions/G1R20_Fix4/tests/run_inherited_excel_regression.py --execute --output <新規旧機能回帰>
```

水理入力表は各ケースの詳細を `tests/water_input_cases.json` に記録しています。逆解析の回数検査は、試験用コピーで宣言した観測カウンターを付け、実際の順解析への入口回数と報告回数を独立に照合します。方程式・収束門は変更しません。

```powershell
python -X utf8 versions/G1R20_Fix4/tests/run_cold_regression_excel.py --case D02A --precision-from-start --budget 600
python -X utf8 validation/G1R19_D02_route_audit_20261016/audit_cold_native.py --run <新規D02A結果> --case D02A --lower <採用下界checkpoint> --upper <採用上界checkpoint> --output <新規監査.json>
python -X utf8 versions/G1R20_Fix4/tests/run_water_native_fixed_excel.py --seconds 240 --solve-seconds 20
python -X utf8 versions/G1R19/native/unified/tests/verify_unified.py --dll versions/G1R20_Fix4/delivery/rpfem_native_x64.dll --output <新規DLL結果>
```

D02Aは入力からの新規解析です。過去の場・メッシュ・Fs・根を読みません。採用したcheckpointは結果summaryと一致するものを指定します。
D08Aは水位表と107要素を新規生成し、q8の下側Fs=1・上側Fs=2を解く固定点の回帰です。Pythonで原q8を独立再組立して場を監査し、根探索・P06本番完走・全26ケース完走と区別します。

Release別添 `EXECUTION_EVIDENCE.zip` に試験コピー・ログ・場・source snapshots・失敗した試行を含めます。通常ZIPには巨大なraw実行証拠を含めません。SHA256と確認範囲は `evidence` を参照してください。
