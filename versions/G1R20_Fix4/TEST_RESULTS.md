# G1R20 Fix4 検証結果

G1R20 Fix3で再現した不具合を修正し、Fix4配布ブックの入力検査と計算・中止・保存移行を回帰検証しました。GAPの標準目標と合格基準は2%です。全26ケースが今回2%で完走したという主張ではありません。

## 配布物との対応

- 配布XLSM SHA256: `b010ab89c4e70540277475d1aab025db2545a24fd6697ad87856998737e94be6`
- source bundle SHA256: `7acc11691df549ea58ed1a063299725f4d5d7f8bc4b2c1fb24c387622ca17a4c`
- 統合DLL SHA256: `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`

配布ブックをcompile・保存・開き直して再compileし、63標準モジュールとThisWorkbookの全64ソースを照合しました。保存XLSMのオフライン抽出もPASSです。source/import128ファイルのUTF8/CP932対応とmanifestを確認しました。Fix3比で14標準モジュールを変更し、49標準モジュールと単一DLLを保持しています。証拠は `delivery/WORKBOOK_VALIDATION.json`、`SOURCE_EXTRACTION_AUDIT.json`、`SOURCE_MANIFEST.json` です。

## Pythonのportable試験

14 suites・132 unittestがPASSしました。過去版の125 testsとFix4の7 testsです。Fix4の幾何試験には矩形・凹形状の独立oracle2,250構成を含みます。これらは入力・ソース契約・幾何判定の試験であり、新規Fs求解やExcel実行の件数ではありません。`evidence/PORTABLE_REGRESSION_FINAL.json` を参照してください。

## 最終Fix4のExcel回帰

|対象|結果|確認した件数と範囲|
|---|---|---|
|主入力5表|PASS|57ケース＝有効11・不正46。ID、参照、物性、部分行、重複領域、容量末尾を含む|
|旧形式テキスト入力|PASS|31ケース＝有効3・不正28。別にbusy状態の保護1。例外時も所有ファイルを解放|
|地震荷重|PASS|17 macro/probe呼び出し。q2/q8、新規・実Fs後・反復、固定/比例、kh0|
|水理入力4モード|PASS|92ケース＝有効22・不正70。各入力契約と選択した小規模solve経路|
|水理中止・γw|PASS|14ケース。公開Stop、逆解析、確定前中止、単位体積重量の有効/不正値|
|水理消費・保存移行|PASS|29ケース。D4/F4・γwの消費拒否、旧proof拒否・新規生成受理、PENDING拒否|
|逆解析の実評価回数|PASS|3ケース。上限6→実6/表示6、12→実12/表示12、最小未満5→0回で拒否|
|設定入力と表示失効|PASS|79ケース・4 sessions。合成TargetMet/表示を用いた操作試験|
|浸透失敗時の状態|PASS|5ケース・5 sessions。公開mode2生産を実行し、合成した土の旧表示・proofの失効、元エラーとアプリ状態復元を確認|
|設定表外キー|PASS|式変化と保存/reopenの2経路＋直接編集control。旧表示失効と公開readerのErr5拒否|
|旧保存結果の移行|PASS|3 sessionsの1シナリオ。物理入力不変、旧表示失効。合成表示のため数値証明ではない|
|配布bookの純公開API|PASS|3 cycles・9直接呼び出し。実Fs→Setup→SeismicSetup。VBEアクセス・製品ソース差替えなし|

水理4 suitesは92＋14＋29＋3＝138ケースです。各suiteの件数を土の本番Fs計算数として数えません。上表のcase、call、cycle、sessionは単位が異なるため総計しません。表示の合成sentinelをFsの数値証明に数えません。

地震の小規模実Fsは下界 `9.999989784787447`、上界 `10.000009993685573` でした。固定地震荷重の積分は `[4,-20,0,0]` で、新規/実Fs後/反復で完全一致しました。支持力は実Fs前後とも `[183.04347822419126,184.89650500619234]` でした。これは履歴依存の回帰であり、斜面benchmark全体の精度証明ではありません。

回数検査では試験コピーのRPX_SeepSolve入口に観測カウンターを1箇所だけ付けました。方程式・物理場・収束門を変更せず、実際の順解析回数と報告回数を独立に比較しました。

証拠は `evidence/NUMERIC_INPUT_EXECUTION_SUMMARY.json`、`evidence/STATE_SETTINGS_FINAL_INDEX.json`、`evidence/WATER_FINAL_RESULTS.json` を参照してください。全runの状態・所有Excel終了・最終source/bookとの対応は `evidence/EXECUTION_RUN_INDEX.json` に記録します。

## 本番入力と旧機能の回帰

|対象|状態|結果と確認範囲|
|---|---|---|
|D02A cold|PASS|236.070秒。下界1.5／上界1.500051451551501、対称GAP0.003430044606%。下界360・上界8要素、両q8。過去の場・メッシュ・Fs・根を読み込まない新規解析。Pythonによるcanonical物理条件・形状・28原CSR証拠・根幅metadata・本番経路traceの監査もPASS|
|旧機能inherited|PASS|プリセット24条件＋実G1R19移行3条件、初期化・例外復旧18件、浸透UI6件、水理source21件、mode2別Excel再起動、保存数式3シナリオ（各2プロセス）。初回のpresetsと再試験7群を組み合わせ全8群を確認。全所有Excel退出、source64不変|
|D08A水連成native固定点|PASS|水位3m・107要素・69節点を新規生成し、下側Fs=1・上側Fs=2を求解。Python独立組立で元E/b/G/h/cost・円錐次元・面積158・鉛直基準合力−2571.4を照合し、両場を原q8で監査。根探索・本番2%完走とは区別|

## 単一DLLの確認

同一SHAの統合DLLに対して、密行列因子化/解120ケース、compiler120ケース、fan/guardと中止100 handles、accuracy FFI11個の小規模fixtureを確認しました。`results/native_abi/summary.json` はPASSです。DLLを変更していないため、これは既存ABIと小規模数値経路の回帰です。26ケースの本番計算完走を置き換える試験ではありません。

## Fix3の再現負例と原記録

公開済みFix3を専用コピーで実行し、元ブック・製品source・DLLの不変を確認しました。Fix3水理92ケースmatrixでは、有効22が通り、不正70のうち入力検査が51を受理し、小規模solveを行った不正例の42が結果資格を得ました。Fix4では同じ92ケースが全て期待通りです。旧版公開Stopは反復1回目で発火した後も258回まで進み結果を公開しました。逆解析も上限6に実7、12に実13を実行していました。これらの負例と他の入力再現は各review、`evidence/WATER_INPUT_NEGATIVE.json`、raw summaryに保存しました。

旧版負例のPASSは「不具合を再現する試験としてPASS」という意味です。最終Fix4の合格件数に足しません。古い候補bookやsourceの証拠も最終版の証拠とは区別します。

初期ハーネスのFAILも保持しています。水理input fixtureの前ケース設定残り、K2空欄/旧PENDINGの判定、期待Err5をPython側で直接扱ったmodal dialogは、それぞれfixture初期化・実consumer確認・試験用catching probeで検証し直しました。製品の合格条件は緩和していません。元FAIL/exit falseを書き換えず、必要な後日退出はexact PID/birthの別観測として保存しています。

継承テスト初回では、無効MODEの拒否メッセージを旧版の日本語文言で期待して1件FAILになりました。実際にはFix4の入力検査が正しくErr5 `RPX_P4 / INVALID_SETTING MODE` を返していました。期待値のみ更新し、初期化18件を含む7群を新たに実行して全てPASSしました。初回の親・子summaryとログも保持しています。`evidence/INHERITED_HARNESS_FIX4_NOTES.md` を参照してください。

## 原因が未確定の異常

Fix3の `results/seismic_negative_/1792183690937194400` で、実Fs成功後の公開メッシュfixture中にExcel回復子プロセスが発生した単発事象があります。直前の内部stageが残っていないため、Setup・SeismicSetup・mesh・DLL・Office描画のどれが原因か確定していません。

追試は3 runsともPASSしました。同じ未改変Fix3の3 cycles/9 calls、quiet5 cycles/15 calls、最終Fix4配布bookの純公開API3 cycles/9 callsで再現しませんでした。これを原因究明済み、DLL修正済み、将来起きないという意味には扱いません。`evidence/SETUP_CRASH_AUDIT.md` を参照してください。

## 確認していないこと

P06A/P06Nの通常初期状態からのVBA/DLL本番GAP2%自動完走、全26ケースのVBA再計算、FOS小数点2桁の参照一致、全入力に不具合が無いことは今回主張しません。GAP1%への追加研究も実施していません。固定点・保存移行・合成表示試験をcold本番完走に読み替えません。

製品VBAの最終独立静的レビューでは追加の高確信不具合は見つかりませんでした。確認範囲は14変更module差分、全64保存ソースとmanifest、class/空sheet code、単一DLLの同一性です。`evidence/FINAL_PEER_REVIEW.json` を参照してください。実行試験や、上記の未解明事象の原因調査を代替するものではありません。

検証ハーネスの最終監査では、4 runnerの後段検証失敗時にsummaryへPASSが残り得る記録上の問題を修正しました。合成24 runner件と12 collector件は132 portable件と別枠で全件PASSです。修正前に完了した正常rawも強化検査を通過し、失敗rawは変更せず保持しています。実Excel証拠は実行時のrunner snapshotを保持し、現行runnerとの差分は記録処理のみです。`evidence/HARNESS_FINAL_CHECK_REVIEW.md` と対応JSONに記録しています。

全Excel試験完了後の記録監査では正常raw29件が通過し、FAIL12件を保持しました（`evidence/FINAL_RECORD_CHECKS.json`）。選択した最終正例16記録は、それぞれ全64ソースのSHAに結び付き、終了・監視・不変性検査も通過しました（`evidence/FINAL_EXECUTION_PROVENANCE.json`）。これは追加の数値求解数ではありません。

証拠ZIPの梱包時には、元記録の終了Falseと後日観測の終了Trueを別項目で保持している形式への参照を修正しました。合成8件で、元FAILの保持、後日観測の受理、不一致・現在生存・重複観測の拒否を確認しました。これは132件・24+12件と別枠です（`evidence/EVIDENCE_PACKAGER_CHECKS.md` と対応JSON）。実証拠ZIPの6,467メンバーは全SHA256・CRC検査を通過しました。

D02AのPython監査は新規VBA/DLL計算が出力した場の独立再監査であり、Pythonの新規求解ではありません。反対側端点のdual/ray全文が未出力の箇所について、Pythonがその符号証明を独立に再証明したとは扱いません。`evidence/D02A_FINAL_PYTHON_AUDIT.json` を参照してください。
