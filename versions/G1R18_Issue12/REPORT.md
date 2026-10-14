# Issue #12 検証報告（2026-10-14）

環境: Triela、Excel 16（COM、非表示、1プロセスずつ）。G1R18配布ブック（sha256 6b280c03…）と修正ブック（5cd46b28…）の別名コピーで実行。設定は ADAPT=1・QUADRATURE=8・RAW_LOWER_AUDIT=1。標準斜面は TARGET=512・CYCLES=1、D07A キャッシュ試験は CYCLES=1、`default` は配布既定（D07A は G1R17 26ケース入力）。

|run|下界|上界|下界q/上界q|両根|キャッシュ命中(最終)|証拠不一致で無効化|最終HSD_NOT_CONVERGED|最終候補拒否|終了|秒|
|---|---:|---:|---|---|---|---|---|---|---|---:|
|base_D07A_acc_cache1|1.2606988771609542|1.2775733308342803|4/4|True/True|1(1)|0|1|1|PARTIAL_CERTIFIED|261|
|base_D07A_default|1.2612500306920593|1.2775733308342803|8/4|True/True|0(0)|0|0|0|PARTIAL_CERTIFIED|157|
|base_std_acc_cache0|1.3951690029037862|1.413119563717958|8/8|True/True|0(0)|0|1|0|PARTIAL_CERTIFIED|206|
|base_std_acc_cache1|1.3943058272230395|1.4132509204501438|4/4|True/True|1(1)|0|2|1|PARTIAL_CERTIFIED|197|
|base_std_default|1.3956340989407385|1.408932402574966|8/8|True/True|0(0)|0|0|0|OPTIMAL|94|
|base_std_fixed_cache0|1.3951690029037862|1.4130621737462|8/8|True/True|0(0)|0|0|0|PARTIAL_CERTIFIED|106|
|base_std_fixed_cache1|1.3943058272230395|1.4132509204501438|4/4|True/True|1(1)|0|1|1|PARTIAL_CERTIFIED|97|
|base_std_off_cache0|1.3951489400304535|1.4130621737462|8/8|True/True|0(0)|0|0|0|PARTIAL_CERTIFIED|103|
|base_std_off_cache1|1.3951489400304535|1.4130621737462|8/8|True/True|1(1)|0|0|0|PARTIAL_CERTIFIED|97|
|fix_D07A_acc_cache0|1.2612567096115295|1.2775733308342803|8/4|True/True|0(0)|0|0|0|PARTIAL_CERTIFIED|272|
|fix_D07A_acc_cache1|1.2612567096115295|1.2775733308342803|8/4|True/True|0(0)|1|0|0|PARTIAL_CERTIFIED|274|
|fix_D07A_default|1.2612500306920593|1.2775733308342803|8/4|True/True|0(0)|0|0|0|PARTIAL_CERTIFIED|156|
|fix_std_acc_cache0|1.3951690029037862|1.413119563717958|8/8|True/True|0(0)|0|1|0|PARTIAL_CERTIFIED|211|
|fix_std_acc_cache1|1.3951690029037862|1.413119563717958|8/8|True/True|0(0)|1|1|0|PARTIAL_CERTIFIED|211|
|fix_std_default|1.3956340989407385|1.408932402574966|8/8|True/True|0(0)|0|0|0|OPTIMAL|94|
|fix_std_fixed_cache0|1.3951690029037862|1.4130621737462|8/8|True/True|0(0)|0|0|0|PARTIAL_CERTIFIED|107|
|fix_std_fixed_cache1|1.3951690029037862|1.4130621737462|8/8|True/True|0(0)|1|0|0|PARTIAL_CERTIFIED|107|
|fix_std_off_cache0|1.3951489400304535|1.4130621737462|8/8|True/True|0(0)|0|0|0|PARTIAL_CERTIFIED|101|
|fix_std_off_cache1|1.3951489400304535|1.4130621737462|8/8|True/True|0(0)|1|0|0|PARTIAL_CERTIFIED|102|

判定:

- G1R18 + キャッシュ1 + 回復経路ON（3件）で Issue を再現: 最終q8でq4キャッシュ命中→負側下界試行が HSD_NOT_CONVERGED→`FINAL_NUMERICAL_UNKNOWN;OLD_WITNESSES_KEPT`、q4保存場に後退。
- 修正版 + キャッシュ1 は全ケースでキャッシュ0と下界・上界が全桁一致し、両根有効、最終候補拒否なし。無効化ログ: `reason=original_evidence_quadrature;cached_q=4;q=8`。
- 両方OFFでは G1R18 でも命中後に正常（ゲート非経由）で、修正版と同値。
- 既定設定の回帰（標準斜面・D07A）は G1R18 と修正版で全桁一致。
- `fHSD=1` の精度回復ケース（キャッシュ0でも同じ）は上界側試行で、Issue と無関係。
- 試験中、ドライバーの入力書込み不具合（D07A）で解析前に失敗した試行があり、修正後に再実行しました（失敗分は解析未実施）。fix_D07A_acc の2ケースは中断復帰で2回ずつ実行され、同値でした。

純Pythonモック: `results/issue12_python_mock.json`（4検査 PASS）。パッケージ: `results/issue12_package.json`（74モジュール、コンパイル・再読込・ソース一致 PASS、他モジュール無変更）。
