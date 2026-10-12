# G1R16 DLL_3

G1R16の数値処理を配列単位でDLLへ移植した64bit Excel向けの版です。通常の内点法とHSDの両方を対象にしています。構成則、Davis変換、Fs定義、許容差、物理監査、再試行方針は維持しています。

## 入手と実行

- [XLSM・DLL・ソース・検証資料のZIP](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R16_DLL_3_20261012.zip)
- [対応XLSM](delivery/RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm)
- [DLL](delivery/rpfem_native_x64.dll)
- [検証報告](REPORT.md) / [実装内容](IMPLEMENTATION_JA.md) / [導入・復旧](IMPORT_ROLLBACK_JA.md)

ZIPを展開し、`delivery/` のXLSMとDLLを同じフォルダーに置いてください。既にDLL版を開いているExcelを終了し、対応XLSMを開きます。「要素定義」の通常の解析操作を使います。例題はTARGET=600、最適化ON、kh=0.15右、浸透連成ONです。

通常の使用にPythonやVBAプロジェクトへのアクセス許可は不要です。DLLが存在しない、読めない、版や必要な関数が合わない場合は、VBAへ自動切替します。古いXLSMにDLLだけ追加する方法ではDLL_3の経路へ入りません。

## 対象処理

| 処理 | DLL_3での移植単位 |
|---|---|
| NTスケーリング | 全コーンの通常精度・DD精度処理 |
| Predictor–corrector補正 | 全ベクトル・全コーンの補正処理 |
| KKT組立 | 基底コピー、コーン寄与、正則化、対角スケーリング、元行列の保存 |
| G積・G転置積 | 全コーンの疎積・加算 |
| HSDの原作用素 | DD精度でのスケール変換、G/H/G転置、等式行列、正則化 |
| HSDの疎LU | 格納、平衡化、行ピボット選択、消去、圧縮、前進・後退代入 |
| LDL・DD-LDL・並べ替え | DLL2の経路を継承 |

Fs探索、求解制御、物理監査、メッシュ・浸透制御はVBAに残しています。元のVBA数値経路も保持し、DLLの有無を切り替えて照合できます。

今回のP06A/P06N検証は97要素・q=2・固定Fs=1の比較です。フルメッシュで最終Fsを再探索した検証や、26ケースの再実行ではありません。詳細は検証報告を参照してください。
