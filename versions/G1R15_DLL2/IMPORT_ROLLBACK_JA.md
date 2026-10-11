# 導入・復旧

配布XLSMは導入済み。ZIP内のXLSMとDLLを同じフォルダーへ置く。
前回DLL版とファイル名が同じDLLを使うため、前回版を開いていたExcelを終了してから開く。
DLLがない場合は、数値核と並べ替えが既存VBAへ自動切替する。

## 元G1R15への手動導入

原本を保全し、別コピーへ導入する。現在の許可済みVBE設定を使用する。

1. RPX_LDL、RPX_SchurLDL、RPX_SeepUI、RPX_Order、RPX_ColumnOrderをエクスポートして保全する。
2. その5標準モジュールを削除し、同名の `import/*.bas` を導入する。
3. 新規標準モジュール `import/RPX_Native.bas` を導入する。
4. 全体をコンパイルして別名保存し、このZIPのDLLを同じフォルダーへ置く。

前回DLL版へ導入する場合も別コピーを使う。RPX_Nativeは追加でなく置換する。
`import/` はCP932・CRLF、往復無損失検証済み。`src/vba/` は編集用UTF-8で、直接VBEインポートしない。
ThisWorkbook・Sheet・クラスモジュールへの配置変更は不要。それらを標準モジュールとして導入しない。

## 復旧

完全復旧は保全した原本、または前回成果物へ戻す。
導入コピーの既存5モジュールを原本へ戻してRPX_Nativeを削除する方法でも元G1R15へ復旧できる。
DLLだけ外す場合はVBAへ戻るが、呼出しフックと前回のプレビュー関数名修正は残る。
DLLを外す・置き換える操作は、そのDLLを利用しているExcelを終了してから行う。

## 再ビルド・試験

`native/build.cmd` で64bit DLLを生成する。Visual Studio Build Toolsのvcvars64.batのパスは環境に合わせる。
数値核のx87式評価・保存境界・演算順を変えない。今回は整数の並べ替えを追加した。
`tests/build_failure_test.cmd` はメモリ確保失敗と未知例外を注入する専用EXEを作り、製品DLLを変更しない。
`tests/NativeTests.bas` は実Excelでの比較用。配布XLSMには残していない。

Pythonツールはpywin32等を必要とする。通常のXLSM解析にPythonは不要。
編集後は `tools/prepare_imports.py` でCP932ファイルを作成する。
旧 `prepare_sources.py` は編集済みソースを初期化するため、この版では実行を拒否する。
`install_and_test.py` は保全済み元G1R15の別コピーへ導入・コンパイル・短い比較試験を行う。
`run_benchmark.py` は別フォルダーで解析し、既存Excelを強制終了しない。
比較・原本ハッシュ検証の一部は、元リポジトリと前回成果物のフォルダー配置を参照する。
実行例・結果は `results/`。ハッシュ一覧はMANIFEST.jsonとSHA256SUMS.txt。
