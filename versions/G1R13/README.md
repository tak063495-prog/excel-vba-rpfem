# G1R13 D03N / D07A

[検証報告](REPORT.md) / [導入・復旧](IMPORT_ROLLBACK_JA.md) / [修正対応](ISSUE_TO_FIX_JA.md) / [全体README](../../README.md)

delivery/の2つのXLSMは共通の修正版ソルバーで、入力と保存結果が異なる。両ケースのFs取得・保存場監査を確認したが、精度目標は未達。REPORT.mdを確認する。

元の納品ZIPをそのまま展開したファイルはMANIFEST.jsonのSHA-256で検査できる。REPORT.md、このREADMEとZIPの検証受領書はGitHub登録時の追記。独立再監査は、このフォルダーでrequirements.txtの環境を準備し、`python -X utf8 tests/verify_portable.py`を実行する。ネイティブExcelの試験スクリプトは元の保全原本やWindows環境への参照を含むため、移行時にそのパスを確認する。登録時にFEM全解析を再実行したという意味ではない。
