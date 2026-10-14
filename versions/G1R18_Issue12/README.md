# G1R18 Issue #12 修正版（下界キャッシュの証拠メタデータ）

[Issue #12](https://github.com/tak063495-prog/excel-vba-rpfem/issues/12) の修正版です。基準は main `56859fc`（G1R18）。G1R18 のコード・配布ブック・DLL は保全し、このフォルダーに修正版を分離しています。

## 利用

- 配布一式ZIP: [`downloads/RPFEM_G1R18_Issue12_20261014.zip`](../../downloads/RPFEM_G1R18_Issue12_20261014.zip)（[SHA-256](../../downloads/SHA256SUMS.txt)）
- `delivery/RPFEM_G1R18Issue12_20261014_1758.xlsm` と同じフォルダーに `rpfem_native_x64.dll`・`rpfem_accuracy_x64.dll`（G1R18 と同一）を置きます。
- 手動適用は **G1R18 の別名コピー**で標準モジュール `RPX_LowerCache` と `RPX_Fs` を `import/`（CP932/CRLF）で置換し、コンパイル・保存します。`src/vba/` は UTF-8 の編集元、差分は `changes.diff`。

## 原因

G1R18 で下界キャッシュ（`LOWER_MODEL_CACHE=1`）に原プログラム情報 `XOriginalCompiled` を保存するようにしました。キャッシュ同一性は q に依存しない下界行列だけで決まるため、粗探索 q4 で保存した `subdivisions=4` が最終 q8 評価で復元され、`RPX_RefreshLower` も strengthFactor しか更新しません。q8 を要求する原双対ゲート・追加DLL回復・固定荷重回復が負側の正常な下界試行を `HSD_NOT_CONVERGED` として拒否し、最終候補が退けられていました。

## 修正（q8 ゲートは維持）

- `RPX_LowerCache`: 保存済み原プログラム情報の q（`RSubdivisions`）または入力世代が現在と異なれば、照合時にキャッシュを無効化（`reason=original_evidence_quadrature`）し、原プログラムを現在の q で再構築させます。復元時にも同条件を再検査（`LOWER_CACHE_EVIDENCE_STALE`）。
- 同一キーの成功時は保存情報を現在の `XOriginalCompiled` に更新し、不整合なら破棄します。新規保存も不整合な情報なら保存しません。失敗後に残ったキャッシュも次の照合で破棄されるため、古い q4 情報は再利用されません。
- `RPX_Fs.RPX_Bound`: キャッシュ以外でも、準備済み下界プログラムの原情報が別の q・入力世代なら再組立します（防御）。
- 監査ゲート（`subdivisions<>8` で拒否）は変更していません。DLL は無変更です。

## 検証（2026-10-14、Triela 実Excel、1プロセスずつ）

詳細は [REPORT.md](REPORT.md)、`results/`。

|モデル・設定|G1R18 キャッシュ1|修正版 キャッシュ1|キャッシュ0（両版同一）|
|---|---|---|---|
|標準斜面・精度回復DLLのみ|最終q8でq4キャッシュ命中→HSD_NOT_CONVERGED、最終候補拒否、q4保存場 1.39430582722304/1.41325092045014|q不一致で無効化→q8再構築、1.39516900290379/1.41311956371796|1.39516900290379/1.41311956371796|
|標準斜面・固定荷重回復のみ|同上 拒否、1.39430582722304/1.41325092045014|1.39516900290379/1.4130621737462|1.39516900290379/1.4130621737462|
|標準斜面・両方OFF|命中・正常 1.39514894003045/1.4130621737462|1.39514894003045/1.4130621737462|1.39514894003045/1.4130621737462|
|D07A・精度回復DLLのみ・CYCLES=1|命中→拒否、q4保存場 1.26069887716095/1.27757333083428|1.26125670961153/1.27757333083428|1.26125670961153/1.27757333083428|

既定設定（`LOWER_MODEL_CACHE=0`）の回帰は G1R18 と全桁一致: 標準斜面 1.39563409894074/1.40893240257497（OPTIMAL）、D07A 1.26125003069206/1.27757333083428（PARTIAL_CERTIFIED）。純 Python 状態モック（Issue の4状態検査、旧版で再現・修正版で解消）も PASS。

## 制限

修正版では最終 q8 の最初の下界で一度だけ再組立が増えます（今回の試験では実行時間差は数秒以内）。この流れでは q8→q8 のキャッシュ命中は発生せず、q8→q8 の通過はモックで確認しました。P06 精度目標（#11）は対象外です。

## 再現

```powershell
python versions/G1R18_Issue12/tests/verify_issue12_mock.py
powershell -NoProfile -File tests/package_issue12_excel.ps1 -BaseBook <G1R18.xlsm> -SrcDir src/vba -OutBook <new.xlsm> -ResultJson <json>
powershell -NoProfile -File tests/run_issue12_matrix.ps1 -Kit <kit> -FixBook <new.xlsm> -Out <dir>
python versions/G1R18_Issue12/tests/summarize_issue12.py <dir> <summary.json>
```

Excel スクリプトは既存 Excel が動作中なら開始しません。自分で起動した PID だけを終了し、元ブックは別名コピーで使用します。
