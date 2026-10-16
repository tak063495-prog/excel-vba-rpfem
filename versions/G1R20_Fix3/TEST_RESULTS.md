# G1R20 Fix3 検証結果

## 対象を固定した確認

- 修正起点: Fix2 main `b1e092ff4a756bd6ccee84af628b494869ed515c`。
- 変更: 7標準モジュールとThisWorkbook。残る56標準モジュールはFix2とbyte一致。
- 配布XLSM SHA256: `09b13bff8d2e01172d8cb46b1065db3887da8e1bfc310cecf5a66c131124d285`。
- source bundle SHA256: `d11f3591e85b6726c5c16b3e7040ccf0718928d15938001d95633d84312518a4`。
- 統合DLL1個のSHA256: `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`（Fix2と同一）。
- 全63標準モジュールとThisWorkbookを組込み、compile→save→reopen→recompileを確認。
  保存XLSMからの独立抽出でも同じソースを確認。入力・Davis方針・GAP2%・q8・根幅条件を保持。
- [最終実行との対応](evidence/POSITIVE_PROVENANCE_FINAL.json): 12個の完了済み正例記録それぞれに、同じ全64コードのSHAが含まれることを照合。

## 再現と修正確認

|試験|結果・内容|
|---|---|
|Portable|125件PASS（Fix3 10件＋継承115件）。数値求解数ではない|
|浸透UI負例|公開Fix2で取消残留3操作と書式更新による18数式の定数化を再現|
|浸透UI正例|6件PASS。Run/check/previewの取消初期化、18数式の保持、旧形式数式移行、Busy中の書式更新拒否|
|保存数式の負例|Fix2保存→別Excel再開後、参照編集で材料41.65→51.65なのに旧表示が残留|
|保存数式の正例|3シナリオPASS。入力一致なら値保持→参照編集で失効、proofなしの未確認注記→編集失効、CELL filename長128→173の再開時変化で即失効。各2所有プロセス|
|ReadModel負例|Fix2で初回TestFs/ExactModelの91と旧制約残留を再現|
|ReadModel/Fs正例|同一/別/再読込などidentity5件PASS。DIRECT_DLL=1維持、q8/raw原監査とOPTIMAL。c10→11→10でFs対9.9999897848–10.0000099937 → 10.9999888219–11.0000109926 → 最初の対を再現|
|水理source|Fix2の公開producer→CreateMeshで4modeの旧水頭使用を再現。Fix3は21項目PASS（lossless 86,016 UTF16コード単位・allmode lifecycle・途中変更・壊れたproof拒否/回復・手動計算式・同一Excel内の4mode save/close/reopen 12条件・公開producer4mode）|
|水理source実プロセス再起動|mode2 PASS。最終配布Fix3コピー、PID12132→22856、イベントOFFでも変更済み旧結果を拒否し、公開producer再生成で正常回復。両退出・全SHA不変|
|継承UI|プリセット24条件＋実G1R19移行3条件、初期化・例外復旧18条件PASS。全所有Excel退出・旧テンプレート不変|
|D02A新規解析|PASS。入力から229.669秒。下界1.5／上界1.500051451551501、対称GAP0.003430044606%（下界360・上界8要素、両q8）。Python原物理・canonical材料/体積/境界荷重・形状・28 original CSR証拠・根幅metadata整合・本番経路trace監査PASS|
|D08A固定点|PASS。新規107要素・69節点、水位3mの連成で下側Fs=1／上側Fs=2。元E/b/G/h/cost/円錐次元、面積158、鉛直基準合力-2571.4をPython独立組立と照合し、両物理場を原q8で監査。場の縮小・Fs変更なし|

通常結果の保存表示試験は合成sentinelを使い、数値証明の作成と区別します。
ReadModelから先のFs誤拒否は、制約初期化だけを修正した途中候補（FsはFix2と同一）で再現し、
値10.000000000000032・original_dual採用でもNO_FINITE_POSITIVE_FSとなることを確認しました。
最終候補は片側ガイドを最適解に格上げせず、解析的Fsの呼出しに既存の最適化経路を使います。

[状態UI集約](evidence/STATE_UI_EXECUTION_SUMMARY.json)、
[ReadModel/Fs集約](evidence/READ_MODEL_FIX3_SUMMARY.json)、
[水理source集約](evidence/WATER_SOURCE_EXECUTION_SUMMARY.json)、
[Portable全結果](evidence/PORTABLE_FINAL.json)、
[D02A独立Python監査](evidence/D02A_FINAL_PYTHON_AUDIT.json)、
[D08A固定点・原q8監査](evidence/D08A_FIXED_FINAL.json)、
[保存ソース監査](SOURCE_EXTRACTION_AUDIT.json)を参照してください。

## 実行履歴の保持

途中候補のPASS、既知不具合を再現したFAIL、fixture/監視条件を調整した試行も削除しません。
旧テキストfixtureの改行、失敗がDLL到達前で未ロードになる期待値、
水理fixtureの必須荷重セルの空欄、Portableの継承パス重複を修正した履歴を保持します。
`PORTABLE_PRELIMINARY.json` と `PORTABLE_DUPLICATE_SUITE_ATTEMPT.json` は最終125件の代わりには使いません。
水理負例初回の終了確認は原記録ではFalseであり、別の時刻付き `POST_RUN_EXIT_CHECK.json` で
元PIDの後日の退出を確認しています。元記録をPASSへ書き換えていません。

Releaseの別添 `EXECUTION_EVIDENCE.zip` に生ログ・使用した試験コピー・原CSR/場・ソース/ハーネスの
スナップショットを収録します。メンバーのSHA256とCRCを検証します。

## 確認範囲

今回のレビューと回帰で不具合が存在しないことを保証するものではありません。
P06の通常初期状態からのVBA本番2%自動完走、全26ケースのVBA再計算は実施していません。
D08Aの固定点監査は根探索や上下界対の本番完走と区別します。
D02AのPython監査は新規VBA/DLL計算が出力した場の独立再監査で、Pythonの新規求解ではありません。
反対側端点のdual/ray全文が未出力の箇所について、Pythonがその符号証明を独立に再証明したとは扱いません。
