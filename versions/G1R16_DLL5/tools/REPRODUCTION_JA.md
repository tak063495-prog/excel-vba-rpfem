# DLL5の検証再現

通常の解析にはPython・コンパイラは不要です。以下は開発・検証用です。試験は専用Excelプロセスと試験コピーで実行します。既存の証跡やユーザーの原本を上書きしない新しい作業場所を使ってください。

## 環境

リポジトリ全体が必要です。旧DLL4/DLL3と`validation/G1R12_native25_20261009/inputs/`、独立監査の`versions/G1R12/validation_runtime/python/`を参照します。ZIPだけから通常解析はできますが、旧版との比較に必要な外部参照すべては含めていません。

Windows・64bit Excel、Pythonのpywin32/numpy/scipy/oletools、ネイティブビルドにはVisual Studioのx64 C/C++・MASMを用意します。`native/build.cmd`と試験用ビルドファイルの`vcvars64.bat`を環境に合わせてください。試験コピーへのモジュール追加には利用者が既に許可したVBAプロジェクトアクセスが必要です。スクリプトはセキュリティ・レジストリを変更しません。

## 保存済み成果の検査

`MANIFEST.json`および`SHA256SUMS.txt`でバイトを検査できます。`results/core_pipeline.json`は22試験グループ、`results/pipeline.json`は8検証手順の終了コードとログの場所を記録しています。`TEST_RESULTS.csv`は実Excel・Python独立検証・ネイティブAPI・コンパイルの別を記載します。

実行の再現は、新しいチェックアウトで`versions/G1R16_DLL5`を`versions/G1R16_DLL5_retest`へコピーして行うと元の証跡を保全できます。`_retest`の`results/`は新規に用意し、元から`baseline.json`、`import_modules.json`、実VBA/DLL4参照の`compiled_q8_first/`だけをコピーしてください。元のフォルダーは削除・変更しません。ソース・インポート・deliveryはそのままコピーします。各ツールは自分の親フォルダーをROOTとするため、以下の例の`G1R16_DLL5`を`G1R16_DLL5_retest`へ読み替えます。

## 主回帰

リポジトリのルートから実行します。

```text
cmd /c versions\G1R16_DLL5\native\build.cmd
python versions/G1R16_DLL5/tools/validate_core.py
```

同じDLLの中断後は`--resume`で終了コード0の試験を保持できます。DLLのSHAが変わった場合は再開を拒否し、新しい最終回帰が必要です。停止した際は`core_pipeline.json`のstate/current、該当`core_N.log`、Excelの実行状態を確認してください。過去の不合格ログを捨てて合格だけにしません。

NT単体比較はExcelの浮動小数点設定を観測し、Pythonの比較スレッドだけを合わせます。製品DLLは設定を変更しません。`test_nt_oracle.py`は必要な観測が未実施なら先に`probe_cw.py`を実行します。

障害試験の出力は実行ごとの別フォルダーへ保存します。`test_new_job_contract.py`はcompiler/Schur/Krylovの開始後へ試験コピーだけの注入点を追加し、成功・取消・メモリ不足・添字・未知例外・events OFF時の入力変更の計24条件を実行します。成功時に正当に保持されるSchurキャッシュは試験終了時に解放してからhandle数を検査します。

## 全解析と限定した難ケース

主回帰がすべて完了してから次を実行します。長時間試験を並行実行しないでください。

```text
python versions/G1R16_DLL5/tools/validation_pipeline.py
```

固定点の全バイナリ照合、DLL5堤防の全解析、D07A/D03Nの粗メッシュ、P06AのFs≥0.125だけ、P06Nの粗メッシュ、数値照合、最終ブックの再読込/コンパイルを逐次実行します。基準DLL4の時間・粗メッシュ比較は保存済み実測参照を使います。今回のDLL4堤防再実測を再現する場合は別途`run_levee.py --backend dll4 --tag levee_dll4_serial`を先に実行します。

標準タグに既に結果がある場合は新しい試験場所を使ってください。`run_case.py`と`run_levee.py`は既存結果への上書きを拒否します。ソース導入用`install.py`を最終結果選択後に実行すると基準から作り直すため、保存した解析結果を置き換えます。導入をやり直す場合は新しいコピーを使ってください。

`run_case.py`はTARGET≤256・適応OFFに制限しています。P06Aには`--minimum-test-fs 0.125`が必須です。この停止点は試験コピーだけの`RPX_Fs.FsValue`に挿入し、製品探索は変えません。`verify_coarse.py`が停止処理以外のソース変更がないことを検査します。最初の15求解までの記録が一致しても、全探索の合格や最終Fs取得ではありません。

P06Nの`--record-lower-bracket-failure`はエラー5 `FS_LOWER_BRACKET_NOT_FOUND`だけを失敗結果として記録します。ドライバー終了コード0は記録の完了であり、statusはfailed、Fs未取得のままです。取消・OOM・添字・未知例外・監査違反をこの扱いに含めません。

フルP06A/P06N、P06A極小Fs、以前中断した26ケース全試験は、このパイプラインでは開始しません。保存済み粗メッシュ結果の再現にはリポジトリの旧DLL4証跡が必要です。独立検証と実Excel数値比較を混同しないでください。

## 配布の検査

`verify_saved.py`はDLL4基準/編集用ソース/実保存VBAを比較し、既存8モジュール変更・1追加・document module無変更、CP932/CRLF、保全ハッシュを確認します。`verify_release.py`は主回帰・固定点・堤防・保存入力/メッシュ・Fsビット一致を必須にし、実Excel保存結果をバイトでコピーして再読込/全コンパイルします。

`report.py`は実行証跡の条件を確認して日本語資料を生成し、`package.py`は原本保全・旧P2-S検査・最終SHAを確認してZIPの全エントリーを検査します。`verify_published.py <commit>`は実際にpushしたcommitと3配布ファイルのHTTP200/サイズ/SHA一致を確認します。GitHubへ送信する操作自体はこれらの検査スクリプトに含めていません。
