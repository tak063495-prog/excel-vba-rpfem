# G1R18_Issue12 — Issue #12 修正候補

q4 の下界キャッシュ証拠メタデータが q8 段へ残り、原双対ゲートで reject される問題を修正します。

## 変更
- `RPX_LowerCache.bas` / `RPX_Fs.bas`：キャッシュヒット時に保存メッシュ段(q)と現在段が異なる証拠を無効化し、q8 で新規に証拠を再計算。
- `src/vba`（UTF-8）、`import`（CP932/CRLF）、`changes.diff`、配布XLSMと既存G1R18 DLL（無変更）を同梱。

## 検証
- Pythonモック `tests/verify_issue12_mock.py`：PASS（旧版は q4 保存・全ゲートFalse、修正版は q8 で全ゲートTrue）。
- 配布ブックのコンパイル・保存・再オープン：PASS。
- 実Excel 14ラン行列（`results/issue12_excel_matrix.md`）：旧版のcache1系は q4/4 で終了し reject 1件、修正版は全ランで q8/8・reject 0。キャッシュ無効時の結果は旧版と一致。
