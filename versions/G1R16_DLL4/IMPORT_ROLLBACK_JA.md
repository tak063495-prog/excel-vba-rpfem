# DLL4の導入・復旧

## 通常の導入

1. Excelを終了します。既に読み込まれたDLLはブックを閉じただけでは交換されない場合があります。
2. ZIPを新しいフォルダーに展開します。
3. `delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm`と`delivery/rpfem_native_x64.dll`を同じフォルダーに置きます。
4. 入手元と`SHA256SUMS.txt`を確認し、通常のマクロ実行手順で対応XLSMを開きます。
5. `RPX_NativeInfo`が`DLL_READY`ならDLL4が有効です。DLL不在や非対応版ではVBAへ切り替わります。

DLLだけをDLL3のXLSMへ追加する方式ではDLL4になりません。元のブックはそのまま保存してください。配布XLSMの初期入力は堤防例題です。

## ソースを取り込む場合

通常は検証済みXLSMを使います。取込みが必要な場合は、DLL3のブックを別名コピーし、`results/import_modules.json`に列挙された標準モジュールを`import/`の同名ファイルに置き換えます。新しい`RPX_NativeSolver`も標準モジュールとして取り込みます。

`src/vba/`は編集用UTF-8、`import/`はCP932・CRLFで無損失往復を確認したVBE用です。`ThisWorkbook`と`Sheet*.cls`は文書モジュールで、新しい標準モジュールとして取り込みません。DLL4は文書モジュールを変更していません。

全プロジェクトをコンパイルし、保存・Excel終了・再起動後に`RPX_NativeInfo`を確認してください。マクロのセキュリティやVBAプロジェクトへのアクセス許可を、この手順で自動変更することはありません。`vbaProject.bin`の入替やスプレッドシートライターによる再構築は行いません。

## 復旧

Excelを終了し、保存しておいたDLL3のXLSMとDLL3のDLLを組で使います。DLL4のブックをVBA経路で使う場合は、Excel終了後にDLLを別の場所へ保管し、ブックを開き直します。解析途中のバックエンド交換は行わず、入力から新しく解析します。
