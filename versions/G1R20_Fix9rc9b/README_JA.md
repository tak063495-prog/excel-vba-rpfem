# G1R20 Fix9 rc9b（候補）

Fix9 rc4b（main）のあとの候補版です。rc5〜rc9a は公開していないため、その変更もすべてこの版に入っています。コードは Codex が書き、こちらで確認しました。

## rc9b で変わったこと（rc9a との違い）

変更したモジュールは `RPX_AccuracyNative.bas`、`RPX_P4.bas`、`RPX_UI.bas` の 3 つです（静的チェック 361 項目 PASS、ほかの 60 モジュールはバイト一致）。

- **上界の元問題を DLL で解くときに平衡化（equilibrate）を使います**（`ACCURACY_DLL_UPPER_EQUILIBRATE`、既定 1。0 にすると rc9a と同じ）。
- 双対の診断ログを直しました（段ごとの記録と呼び出し番号）。判定と許容差は変えていません。
- 検討した「状態 5 で早めに棄却」は入れていません。実際の上界を捨ててしまう例が 40 件見つかったためです（`codex/st5scan.txt`）。

## rc9a の内容（この版に含まれます）

- 大きい上界の段で、解が得られないまま時間切れになったときの VBA エラー 9 を直しました。
- 後の段でエラーが起きても、確定済みの下界・上界と Fs を書き出します。
- 詳しくは `codex/rc9a/` にあります。

## rc5〜rc8 の内容（要点）

- 薄層チェーンの DLL 解法で平衡化（rc8、`ADAPT_THIN_DLL_EQUILIBRATE`）。下界の追加段は既定 3 段です。
- 同じメッシュ上の Fs 試行のあいだで HSD の状態を引き継ぎます（rc7b）。
- 大きい上界の HSD 時間・反復の拡大、根探索の区間の絞り込み、上界の分割順の見直し（rc7）。薄層チェーンの既定時計 18171 秒（rc6c）。上界の段の要素数制限など（rc6）。状態 5 の扱い、最良解の保持など（rc5）。
- rc4b との差分（`src/`）は `docs/diff_rc4b_to_rc9b.txt`（行末の違いは無視）と `docs/diff_summary_rc4b_to_rc9b.txt` にあります。

## 配布物（`delivery/`）

| ファイル | 内容 |
|---|---|
| `RPFEM_20261021_G1R20_Fix9rc9b.xlsm` | ブック（コンパイル・開き直しの確認 PASS）。SHA256 `f025cb2acac0e261f2333728677f96cd612d2de4adac30f2eac9f440a65fd788` |
| `rpfem_native_x64.dll` | DLL（Fix9 と同じもの、SHA256 `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`） |
| `WORKBOOK_VALIDATION.json` | コンパイル・開き直しの確認記録 |

ブックと DLL は同じフォルダに置いてください。ハッシュは `SHA256SUMS.txt` にあります。ブックの「実行案内」シートには、作成時の文言「未公開・未検証」が残っています。

## 確かめたこと（Winmax2、日本時間）

両ケースとも追加の設定は `ADAPT_THIN_UPPER_ROUND_MAX_ELEMENTS=8000`、`ADAPT_THIN_UPPER_CAP=9000`、`ADAPT_THIN_EARLY_UPPER_CAP=9000`、`ACCURACY_DLL_UPPER_EQUILIBRATE=1`、制限時間は各 18171 秒です。

| ケース | 状態 | Fs（下界〜上界） | GAP | 時間 | rc8 |
|---|---|---|---:|---:|---|
| P06A（2026-10-22 02:15 終了） | PASS_TARGET | 0.84307764（1481 要素）〜0.85815943（2944 要素） | 1.773% | 7470 秒 | 1.770%、7151 秒 |
| P06N（2026-10-22 00:10 終了） | FS_AVAILABLE_PARTIAL | 0.67486963（2000 要素）〜0.69145410（2938 要素） | 2.428% | 17690 秒 | 2.428%（Fs はビット一致）、13552 秒 |

- **P06A は 2% を下回りました**（独立監査 PASS、ソース・DLL のハッシュは実行前後で同じ）。
- **P06N は 2% に届いていません。** 上界の平衡化で上界の段 0〜7 の HSD への切り替えはなくなりましたが、Fs は変わりませんでした。4700 要素の上界の段は制限時間に達しました。下界は rc8 より遅くなっています。監査は PASS です。
- **7 ケースのセットは rc9b では実行していません。**
- 気になる点（未確認）：P06N のログに `thin_layer_step seconds=-84046` があります。VBA の `Timer` が深夜 0 時に戻ったためと思われます。時間保護の判定に影響するかは確かめていません。
- 短い再現（P02A_THIN）では、平衡化あり（既定）が GAP 0.867%、なし が 0.723% で、このケースでは悪くなっています（`TEST_RESULTS.md`）。

検証の記録は `validation/G1R20_Fix9rc9b_p06n_20261021/` と `validation/G1R20_Fix9rc9b_p06a_20261021/` にあります。

## 元に戻すには

- `ACCURACY_DLL_UPPER_EQUILIBRATE=0` で rc9a と同じ動きになります。

## そのほか

- `tools/` はブックの作成用です。`tools/build_ps51.ps1` と `docs/BUILD_DLL.md` は DLL の作り直しの手順です（Winmax2 で、変更なしのソースから作り直して機械語が同じことを確認済み）。`build_ps51.ps1` は本来 `versions/G1R19/native/unified/` に置いて使うものです。
- `codex/` は Codex の設計・テスト計画・静的チェックの記録です。
