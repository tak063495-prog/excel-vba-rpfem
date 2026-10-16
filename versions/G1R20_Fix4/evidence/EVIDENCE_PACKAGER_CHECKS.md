# 証拠ZIPの後日終了観測schema確認

最初の証拠packagerは、既存の `settings_extent_positive_01/POST_RUN_EXIT_CHECK.json` に保存された `original_owner_exited=False` を後日観測の判定として読んで停止した。ZIPはまだ作成されていなかった。この項目は元FAIL時点の終了flagを保存したもので、後日観測は `post_run_owned_excel_exited=True`、`owned_excel_exited=True`、`process_present=False` に記録されている。原summaryはFAIL、原終了flagはFalseのまま保持する。

packagerは明示された `post_run_owned_excel_exited` を優先し、Falseなら旧flagがTrueでも拒否する。`process_present` があるschemaではFalseを要求する。元のPID・作成時刻と一意に一致する後日観測、観測日時、現在の同一PID・作成時刻のプロセス不在も引き続き検査する。PIDが再利用されても作成時刻が異なれば元ownerは終了済みと判定できる。プロセスの終了操作は行わない。

`tests/run_evidence_packager_checks.py` は専用の新規scratch下に合成rawを作り、ROOT/REPOだけを差し替えた実packagerの `main()` を8回呼ぶ。受理する2例は原False＋後日Trueと旧単一観測schema。拒否する6例は後日False、PID不一致、作成時刻不一致、現実行Pythonの同一PID・作成時刻が生存中、`process_present=True`、重複観測である。受理時にはZIPのCRC・原FAILと原終了Falseの保持・rawのbyte一致を検査し、拒否時にはZIP未作成を検査する。

**8件全てPASS**。Excel/solver実行は0、終了したプロセスも0。製品63標準モジュール・ThisWorkbook・配布ブック・DLL・manifest・packagerとテスト自身・既存raw summaryは不変だった。結果とhashは `EVIDENCE_PACKAGER_CHECKS.json` に記録した。この8件は132 portable件、および後段記録の24＋12件とは別枠である。

```powershell
python -X utf8 versions/G1R20_Fix4/tests/run_evidence_packager_checks.py --output <新規packager確認.json>
```

試験はWindows/Linuxで実行可能で、実Excelを必要としない。既存FAILをpositiveへ変更する検査ではなく、失敗証拠をそのまま安全に保存できることを確認する。

## 実証拠の再梱包

初回停止ログ `D:\Work\rpfem_fix4_review_20261017\evidence_packager.log` は保持した。SHA256は `9f2c30bd6ff9b7aaa68570863780e7586822a8ae5c6741a3d195f8fdf0020a74`。停止箇所は旧フィールドを読むassertで、ZIP作成処理より前だった。

修正後、42個の実結果rootからの再梱包はPASS。`EXECUTION_EVIDENCE.zip` は346,652,274 bytes、6,467 membersで、全memberのSHA256とCRCを照合した。ZIP SHA256は `0008756f1a598f0f3913faea4c759ebb50e17a01c12e7945e0d9f9351fce6d1a`。成功ログ `D:\Work\rpfem_fix4_review_20261017\evidence_packager_02.log` のSHA256は `4ffedbaf95fc8dc140ceec3f70d13fe2c480f74f2783a2c4fb736a7433050572`。要約は既存の `EXECUTION_EVIDENCE.json` に記録されている。
