# DLL_3の導入と復旧

## 完成済みブックを使用する場合

1. 作業中のExcelブックを保存し、DLL版を開いているExcelを終了します。
2. ZIPの`delivery/`を別の作業フォルダーへ展開します。
3. `RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm`と`rpfem_native_x64.dll`を同じフォルダーに置きます。
4. 入手元と`SHA256SUMS.txt`を確認し、通常のマクロ実行手順でXLSMを開きます。

この版のDLL ABIは`0x00160003`です。DLL2のABIとは異なります。旧DLLを混在させるとVBAへ切り替わります。

## ソースを取り込む場合

対象はG1R16 DLL2の対応ブックです。まず別名コピーとバックアップを作ります。UTF-8の`src/vba/`は編集用です。VBEへの取込にはCP932・CRLF・往復一致を確認した`import/`を使用します。

| 既存モジュールの差替え | 新規標準モジュール |
|---|---|
| RPX_Conic、RPX_HSDLu、RPX_Lower、RPX_LowerCache、RPX_Native、RPX_Solve、RPX_Upper | RPX_NativeBulk、RPX_NativeLU |

既存標準モジュールを対応する取込ファイルへ差し替え、モジュール名を維持します。`ThisWorkbook`と`Sheet*.cls`は変更していないため、差替えも新規標準モジュールとしての取込も不要です。全体をコンパイルし、同じフォルダーへDLLを置きます。検証用の`tests/`モジュールは通常のブックへ取り込まないでください。

32bit Excelの実行試験は今回未実施です。DLL経路は64bit用です。コード上の32bit経路はVBAへ切り替わる構成です。

## 状態の確認

必要な場合、VBEのイミディエイトウィンドウで`? RPX_NativeInfo()`を実行できます。`DLL_READY`がDLL使用可能、`DLL_NOT_FOUND`や`DLL_ENTRY_MISSING`などがVBAへの切替理由です。`nt_bulk`、`hsd_operator`、`hsd_lu`などに実行回数を表示します。使われなかった経路の回数が0なのは正常です。

比較用には`RPX_NativeSetMode "vba"`、自動選択に戻すには`RPX_NativeSetMode "auto"`を使用します。求解中に変更せず、次の解析開始前に指定します。

## 復旧

元のG1R16ブックとDLL2を別フォルダーで開けば元の版へ戻れます。DLL_3ブックでDLLを外した場合もVBAで計算できます。DLLを交換する際は、読込済みDLLが残らないようExcelを終了してから交換します。

この実装作業ではVBA信頼アクセス、マクロセキュリティ、レジストリを変更していません。XLSMの保存と取込には実Excelを使用し、`vbaProject.bin`の交換やスプレッドシート書込ライブラリによる再構築は行っていません。
