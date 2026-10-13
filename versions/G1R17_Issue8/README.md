# G1R17 Issue #8 修正候補

[Issue #8](https://github.com/tak063495-prog/excel-vba-rpfem/issues/8) の静水圧誤判定を修正したレビュー用ブランチです。基準は main `5538cb0bf0a47d8adfc50fcd6d9cfcd6faf54a60`。既存 G1R17 のコード・配布ブック・DLL・マニフェストは保全し、このフォルダーに修正候補を分離しています。正式リリース・main への反映は未実施です。

## 利用

`delivery/` の別名 P06A/P06N XLSM と同じフォルダーの DLL を組で使用します。DLL は元の G1R17 配布版と同一です。元のブックと解析データを保管し、別名コピーで入力して再解析してください。保存済みの P06A/P06N Fs は以前の粗メッシュ結果で、今回 Fs を再計算した値ではありません。

手動適用する場合は **G1R17 の別名コピー**で、標準モジュール `RPX_Seep` と `RPX_SeepUI` を `import/` の CP932/CRLF ファイルで置換します。コンパイル、保存、再読込を確認します。`src/vba/` は UTF-8 の編集元です。ThisWorkbook・Sheet モジュールへの追加はありません。マクロや VBA アクセスの設定変更・ソフト導入は行っていません。

## 変更

- 自由節点の残差を `SResTol * Q + gamma * A_i` と比較。`A_i = Σ|Kij^e hj| + |指定流量|`、`gamma = u/(1-u)`、`u=(項数+32)*2^-52`。ゼロの許容値では残差ゼロだけを受理します。
- 緩和前水頭差 `dfix < STol` と浸出面切替なしの条件を維持。小さい緩和係数による #6 の誤収束へ戻しません。
- **SLastResidual は「残差／混合許容値」、閾値は 1 に変更**。SResTol は境界流量に対する相対係数のままです。外部利用側が旧閾値と比較している場合は更新が必要です。
- 結果シート W1:X7 と未収束メッセージを更新。生の最大自由残差、境界流量尺度、最大丸め許容値、丸め成分が相対成分を超える節点の有無も出力します。最大残差と最大丸め許容値は同じ節点とは限りません。

## 実施済み検証（2026-10-13）

|項目|結果|
|---|---|
|標準 Python 数値ミラー|全 assert PASS。静水圧旧300回 False/res=8 → 新1回 True、rho≈0.009795918|
|実 Excel 16.0 静水圧|旧300回 False/res=16 → 新1回 True、水頭=[10,10,10]|
|Excel 純 VBA / DLL_READY|両モードで静水圧54変種/モード PASS（1/2要素、透水尺度3種、水頭3種、CCW巡回3種）|
|#6 極小緩和 1e-8|両モード300回未収束。dfix≈0.66666467、rho≈333332.4。結果シート経由の連成拒否を維持|
|非線形ダム 12×6・20×10・40×20|旧版と新判定で全水頭・反復回数一致、両モードで同じ結果。標準40×20は54回、Q=2.4181|
|修正ブック P06A/P06N|71モジュール、コンパイル/保存/再読込 PASS、修正2モジュールのソース一致、他はVBEの識別子ケース統一だけ。保存ブック再読込後に静水圧1回収束|
|DLL 再ビルド|隔離コピーで成功。ABI 0x00170005、94エクスポート一致、既存疎コンパイラPython回帰120件 PASS|

Excel の浸透ソルバーは自己完結した VBA skyline/RCM 実装で、DLL を呼びません。`DLL_READY` の試験は DLL 利用ブック内で同じ浸透ソルバーを実行した確認であり、DLL 独自の浸透判定の検証を意味しません。DLL C/C++/ASM ソースは全件照合し無変更です。再ビルド DLL は試験用で、配布 DLL を置換していません。

Excel 試験は新規 COM インスタンスの PID が既存プロセスでないことを確認して実施。既存 Excel・作業ブック・解析には接続・停止・変更していません。

## 再現

リポジトリルートから:

```powershell
python versions/G1R17_Issue8/tests/verify_issue8.py
powershell -NoProfile -File versions/G1R17_Issue8/tests/run_issue8_excel.ps1
powershell -NoProfile -File versions/G1R17_Issue8/tests/package_issue8_excel.ps1
```

Excel スクリプトは既存の Excel/VBA アクセス許可を必要とします。許可がない場合は失敗として止まり、設定を変更しません。試験の入力元は元 G1R17 配布ブックです。再パッケージ時には候補ブックを再生成します。

DLL は `../G1R17/native/build.cmd` が公式手順です。この環境にはその固定パスの VS18 BuildTools がなく、VS18 Professional と VS2022 BuildTools がありました。隔離ディレクトリ `issue8-build/native` に native 全体をコピーし、build.cmd の vcvars64.bat パスだけを `C:\Program Files\Microsoft Visual Studio\18\Professional\VC\Auxiliary\Build\vcvars64.bat` に合わせました。`ml64`、C11/C++17、`/O2 /fp:strict /MT /W4 /WX`、リンクライブラリは公式手順のままです。コンパイラは MSVC 14.51.36231。元 native と delivery は変更していません。

```powershell
cmd /c issue8-build\native\build.cmd
python versions/G1R17_Issue8/tests/verify_issue8_native.py issue8-build/delivery/rpfem_native_x64.dll
powershell -NoProfile -File versions/G1R17_Issue8/tests/run_issue8_excel.ps1 -RebuiltDll issue8-build/delivery/rpfem_native_x64.dll
```

native 回帰の既存 oracle は NumPy を使用します。Issue #8 の数値ミラーは標準 Python だけで動作します。GitHub Actions はこの数値ミラーを実行し、Excel/DLL の実行は含みません。

## 制限

丸め係数32は工学的な許容枠で、座標の桁落ち、極端な歪み、Exp、線形ソルバー誤差を含む厳密上界ではありません。極小流量・大きい水頭基準では丸め成分が支配し、相対精度1e-6を保証しません。必要なら水頭基準のシフトなどを別途検討してください。

浸出面の丸め切替振動は別課題です。Python の既知例は旧・新とも20回切替で未収束を維持します。今回、大規模・材料境界・極端な悪条件要素、連成 Fs の全解析、#1/#2 の全回帰は再実行していません。既存の粗メッシュ証拠やベンチマーク値を今回の測定として扱いません。

詳細は `results/issue8_python.json`、`issue8_excel.json`、`issue8_package.json`、`issue8_native.json`、`issue8_native_compile.json`、`issue8_build.log`。配布・ソースの現在ハッシュは本フォルダーの `MANIFEST.json` と `SHA256SUMS.txt` に記録します。
