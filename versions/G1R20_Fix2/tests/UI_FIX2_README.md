# Fix2 UI / Control regression

`run_ui_setup_control_excel.py` は専有Excelを作り、18条件を別ブックのコピーで確認します。数値場・根・解ステータスの移植やSOCP/Fs求解はしません。Currentフラグは状態遷移を検査するための試験用sentinelです。メッシュ生成の正常/異常試験だけは目標16の実メッシュ生成を行います。

| 条件 | 件数 | 検査する結果 |
|---|---:|---|
| Setup: 入力保護、正常、Busy × 外側InternalWrite 0/1 | 6 | Busy時は入力不変で拒否。それ以外は旧メッシュ/結果を無効化し、成功/失敗の両方で元のInternalWriteを復元 |
| CreateMesh: 開始前の図保護、外側write中の結果保護、開始後の節点保護、正常 | 4 | 元の失敗はLastFailureに保持。作業開始前はCancelKeyに触れず、開始後の二重エラーでもBusy解除・CancelKey復元 |
| 3 preset: GAP欠落＋結果保護 × 外側InternalWrite 0/1 | 6 | 二次1004で元error5/RPX_UI/ACCURACY_SETTING_MISSING GAPを失わない。外側のwrite scopeを保持 |
| Run: 結果保護、不正MODE | 2 | 求解前の元エラーを保持し、Busy解除・ResultCurrent=False・CancelKey復元。既存のMeshCurrent保存仕様は評価対象外 |

Run/CreateMeshは従来から失敗を外へ再送出する契約ではありません。正常に後処理へ戻りつつ、元のエラー番号・Source・説明を`RPX_LastFailureNumber`、`RPX_LastFailureSource`、`RPX_LastError`に残すことを確認します。失敗を成功へ読み替える試験ではありません。

リポジトリのルートで実行します。既存Excelプロセスがある場合は中止し、出力フォルダの上書きも拒否します。

```powershell
python -X utf8 versions/G1R20_Fix2/tests/run_ui_setup_control_excel.py --execute --published-only --output versions/G1R20_Fix2/results/UI_negative_fresh
python -X utf8 versions/G1R20_Fix2/tests/run_ui_setup_control_excel.py --execute --output versions/G1R20_Fix2/results/UI_fixed_fresh
python -X utf8 versions/G1R20_Fix2/tests/UI_run_excel.py --execute --output versions/G1R20_Fix2/results/UI_presets_migration_fresh
python -X utf8 -m unittest discover -s versions/G1R20_Fix2/tests -p test_fix2_ui_contract.py -v
```

負例は公開Fix1ブックへ試験用モジュールだけを加えます。正例は同じコピーへFix2の全63標準モジュールを読み込み、ロード内容を照合・コンパイルします。元配布ブックを保存しません。全63ソースの前後SHA、元ブックSHA、起動PID・作成時刻・完全コマンド、退出、ダイアログをsummaryへ記録します。240秒の停止監視はPID・作成時刻・コマンドが一致する専有プロセスだけを対象にします。

追加の`UI_run_excel.py`はFix1から継承した24状態条件と公開G19ブックからの3移行条件です。Setupを呼ばず、G19で不足する7設定が作成されること、物理入力・式・解析政策が変わらないことを確認します。

中間試行も保存します。`UI_setup_control_fixed_01`は最初の修正案がactive error handler内の二次1004に耐えないことを検出した製品側の失敗です。`fixed_02`はその修正後の当初10条件の合格です。追加負例`negative_02`ではRun失敗時にMeshCurrent=Falseを要求する試験側の誤った仮定が1条件あり、最終ハーネスでは従来の保存仕様を維持してその値を評価対象から外しました。`fixed_03`は18状態条件と退出が合格しましたが、試験用監視スレッドが閉じたウィンドウの1400エラーで停止したため、監視も含む最終合格へ加算しません。最終ハーネスは1400だけを記録して監視を継続し、他の監視例外は失敗扱いとします。負例`negative_03`と最終正例は同じ18状態条件を使います。これらの中間試行を最終合格数へ加算しません。

ソース契約ではSetupの既定入力と設定値、3 presetの設定本文、Runの数値/失敗本文の一致を検査します。Runと3 presetの変更は、元エラーを保存した後に`Resume FailedCleanup`でactive handlerを終えて後処理へ入るための2行だけです。その他のUI/Control手続きも指定した7手続き以外は一致を要求します。
