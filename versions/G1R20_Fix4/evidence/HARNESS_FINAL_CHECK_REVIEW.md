# 実行記録の最終確認（harnessのみ）

入力・状態・移行の4 runnerには、正常ケース終了後に `status=PASS` を保存し、その後で所有Excelの終了・監視エラー・ダイアログ・timeout・監視ファイル不変をassertする経路があった。後段assertが失敗するとプロセスは非0終了するが、raw summaryにはPASSが残り得た。旧assertを使った合成monitor errorでこの記録上の問題を再現した。製品VBA・数式・DLLの不具合ではない。

`run_saved_revision_fix4.py`、`run_settings_input_fix4.py`、`run_seep_state_fix4.py`、`run_settings_extent_display_fix4.py` は、後段検証に失敗した場合にFAILと理由を保存してから例外を返す。`collect_release_validation.py` も、root/session/stage/runに記録された監視・cleanupエラー、ダイアログ、timeout、ファイル変更、非0終了を拒否する。入力負例で期待されるエラーは `cases` の結果であり、これらの実行環境エラーとは区別する。

再実行可能な `tests/run_final_record_checks.py` はExcelを起動せず、runner末尾の検証コードをASTから取り出して合成記録に適用する。runner **24件** とcollector **12件** が全件PASS。これらは132件のportable契約とは別枠で、portableの件数を増やしていない。製品ソース・配布ブック・DLLの不変も確認した。

`HARNESS_FINAL_CHECK_REVIEW.json` 作成時に完了済みのPASS raw **18件** は、全て強化collectorを通過し、後段検証との矛盾はなかった。FAIL raw **10件** はFAILのまま保持した。既存raw summaryは書き換えていない。実行中のD02記録はこの完了済み監査に含めていない。

4 runnerの現行SHA256は既存Excel試験時の `runner_snapshot.py` と異なる。JSONの `historical_harness_snapshots` に両方のhashを明記した。差分は後段失敗時の記録処理のみで、既存正常ケースのfixture・呼出し・判定・製品コードを変更していない。既存の正常なExcel実行証拠を、新runnerで実行したものとして扱わない。今回の後段検証に関する合成試験と既存rawの再監査を追加証拠とし、製品の数値再計算は別途の実行記録に依拠する。

```powershell
# 新しい出力ファイルを指定する。--audit-records はclone内の完了済みrawも読み取り監査する。
python -X utf8 versions/G1R20_Fix4/tests/run_final_record_checks.py --output <新規harness確認.json> --audit-records
```

通常配布ZIPだけで合成24+12件を再実行する場合は `--audit-records` を省く。raw実行証拠は別添 `EXECUTION_EVIDENCE.zip` に保存する。失敗rawの後日終了確認は原FAILを変更せず、別の `POST_RUN_EXIT_CHECK.json` として扱い、positive集約へは採用しない。
