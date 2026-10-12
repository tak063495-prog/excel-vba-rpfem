# DLL4の検証再現

通常の解析にはPythonやコンパイラは不要です。以下は開発・検証用です。ユーザーの開いているExcelは操作せず、各スクリプトは専用のExcelプロセスと試験コピーを使います。

## 環境

- リポジトリ全体を取得します。`versions/G1R16_DLL_3/`と`validation/G1R12_native25_20261009/inputs/`も必要です。ZIPだけを別の場所へ展開した場合、実行用XLSM/DLLは使えますが、比較用旧版とケースカタログは含まれません。
- Windows、64bit Excel、Pythonのpywin32・numpy・scipy・oletoolsを用意します。
- Native C/C++試験・ビルドにはVS Build Toolsのx64 C/C++・MASMが必要です。`native/build.cmd`と`tests/build*_failure_test.cmd`の`vcvars64.bat`パスは実際の環境に合わせます。
- 既存のVBAプロジェクトアクセス許可が必要な試験があります。スクリプトはマクロのセキュリティや許可・レジストリを自動変更しません。
- 本番のブックを置き換える前に、原本とDLL_3の組を保管します。

## 手順

リポジトリのルートから、使用中のPythonで次を実行します。既存の証跡を保全する場合は別のチェックアウトと新しい試験タグを使います。`--tag`には英数字・アンダースコアを使います。

```text
cmd /c versions\G1R16_DLL4\native\build.cmd
python versions/G1R16_DLL4/tools/test_native_faults.py
python versions/G1R16_DLL4/tools/test_job_lifecycle.py
python versions/G1R16_DLL4/tools/test_hsd_regression.py
python versions/G1R16_DLL4/tools/test_job_contract.py
python versions/G1R16_DLL4/tools/test_mode_switch.py --tag mode_switch_recheck
python versions/G1R16_DLL4/tools/test_lu_oracle.py
python versions/G1R16_DLL4/tools/test_bulk_oracle.py
python versions/G1R16_DLL4/tools/test_boundary.py
python versions/G1R16_DLL4/tools/test_issue_regression.py
python versions/G1R16_DLL4/tools/test_fallback_seep.py
python versions/G1R16_DLL4/tools/benchmark_fixed.py --tag fixed_q8_recheck --target 256 --q 8
python versions/G1R16_DLL4/tools/run_levee.py --backend dll3 --tag levee3_recheck
python versions/G1R16_DLL4/tools/run_levee.py --backend dll4 --tag levee4_recheck
python versions/G1R16_DLL4/tools/run_case.py P06N --backend dll4 --tag p06n_recheck --target 200 --q 8 --adapt 0
```

NTのctypes単体比較は、実Excelのx87制御語を試験プロセス側に合わせる必要があります。収録の`results/nt_oracle/controlword.json`は、このPCの実Excelで取得した値（DLL_3の証跡から継承）です。別環境では実測値を用いてください。これは試験用プロセスの設定であり、製品DLLがExcelの制御語を変更する処理ではありません。

`test_nt_oracle.py --extended`は実ExcelのVBAオラクル224入力を比較します。`test_lu_oracle.py`は検証用だけ因子を取り出します。製品のLU求解は因子をVBAへ戻しません。`test_job_contract.py`は試験コピーへ故障注入を加えるため、そこで作るXLSMを製品用に使わないでください。

`install.py`は、DLL_3の別名コピーに変更済み7モジュールを実Excelで取り込み、配布候補を保存する**開発用処理**です。配布用の計算結果ブックを保持したい場合に再実行すると、初期ブックへ戻るため注意してください。`prepare.py`は初回作業領域準備用で、再実行は既存ソースを保護するため拒否します。

## 結果の区別

`PASS`は各スクリプトの明示的な検証条件が実行済みであることを示します。固定Fsの比較で入力したFsを最終安全率として扱わないでください。`run_case.py`は実ログの`OPTIMAL`と`PARTIAL_CERTIFIED`を区別し、後者は目標区間幅の未達として記録します。P06Aフル試験はユーザー指示で約66分後に`RPX_Stop`により中断し、実ログは`CANCELLED`です。旧ハーネスの原記録`failed`と取消証跡を残し、`root_regressions.json`で実際の取消として区別しています。P06Nフル試験は開始していません。

速度比較は他の長時間試験と同時に実行せず、同じ入力・メッシュ・Fs点で行います。`validation_pipeline.py`は今回のP06A終了を待ってから、固定点・堤防・粗メッシュFs探索を順次実行した最初の手順です。P06Nの粗メッシュで基準版が正の下界を囲めず、最初のハーネスは停止しました。その原記録を残し、`resume_comparisons.py`で失敗結果も同じ条件で比較します。

`run_case.py --record-lower-bracket-failure`は元のエラー5 `FS_LOWER_BRACKET_NOT_FOUND`だけを「実行を記録した失敗ケース」として扱います。終了コード0は比較ドライバーの記録完了を表し、ケースの解析合格ではありません。`status.json`の`status`は`failed`、`Fs_obtained`はfalseのままです。取消・メモリ不足・添字・監査・未知例外はこの扱いに含みません。`verify_release.py`で両版の探索点・数値・監査・終端エラーの一致を別に判定します。

P06Aの基準DLL_3は粗メッシュの23求解目、Fs=0.00048828125で`HSD_NOT_CONVERGED`となりました。ユーザーは極小Fsの追加探索を終了し、取得済みのFs≥0.125の15点を照合して仕上げる範囲を選択しました。`finish_bounded_comparisons.py`はこの範囲でDLL4を実行し、その後D03N/D07Aの両版を順番に実行した手順です。既存の結果を保全するため、そのまま再実行するスクリプトではありません。

このP06A試験には`run_case.py --minimum-test-fs 0.125`を使用しました。試験コピーの`RPX_Fs.FsValue`入口だけで、下界Fs<0.125になる前にエラー18 `TEST_SCOPE_MIN_FS_FLOOR`を送ります。製品ブックの探索制御は変更しません。`verify_release.py`はコピーの変更がこの停止処理だけであることと、最初の15点の求解・監査・入力・メッシュの一致を検証します。時間も両版の15点目の完了時点で比較します。`BOUNDED_TEST_STOP`は試験範囲の終了であり、解析の完走や最終Fs取得ではありません。基準版の23点目の未収束をDLL4で再現する試験は`NOT_RUN`です。

別の試験タグでの再現には対象フォルダー名を合わせてください。全ケースのフル適応計算を自動で合格にする処理ではありません。
