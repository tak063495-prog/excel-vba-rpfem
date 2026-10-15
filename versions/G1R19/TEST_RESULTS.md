# G1R19 検証結果（2026-10-15）

最終単一DLL、VBAの採用・復帰と保存場再利用、水理・キャッシュ、配布ブックの検証を通過しました。
本番規模のP06A/P06N保存済み4場も、元のVBA q8式と独立再計算で合格しました。
1%を狙う追加実験は停止し、合格基準は対称上下界幅2%以下としています。

## 単一DLLとネイティブ回帰

- 最終DLL SHA-256: `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`
- 従来94＋精度回復9＝103 exports。両ABIを保持し、依存はWindows標準DLLのみ。
- Rust 4件、q8/worker/LSMRのFFI 17件、旧LDL/DD 120ケースのビット一致、sparse compiler 120ケース。
- fan回帰9項目と取消100ハンドル。元の262入力ファイルと統合用268コピーのSHAを照合。
- [ローカル回帰の原結果](https://github.com/tak063495-prog/excel-vba-rpfem/blob/main/versions/G1R19/native/unified/validation/final_g1r19/summary.json) / [最終ビルド証拠](https://github.com/tak063495-prog/excel-vba-rpfem/blob/main/versions/G1R19/native/unified/validation/final_verification_manifest.json)。
- GitHubは固定Rustからの再ビルドDLLと配布DLLを別々に検査し、[PR #17のchecks](https://github.com/tak063495-prog/excel-vba-rpfem/pull/17/checks)で結果を公開します。水理のPython回帰も確認対象です。

## 実Excelの制御回帰

- 最終ソース・最終DLLで固定Fs 9例、候補棄却・強制DLL障害からの既存経路への復帰を確認。
- 採用した原primal/原full-dualの全ベクトルを、元CSR全行・cone・原目的で独立再計算。
- 同一実行q8保存場の最終再利用は正常2場、不一致・場破損16例、追加根探索0回。場のバイトと既存上下界を保持。
- この制御fixtureのGAP=1.9は製品2%達成の証拠とは区別。製品判定は別unitで対称幅1.99%を採用、2.01%を拒否し、不完全ペアも拒否。
- q4キャッシュからq8への破棄・再構築と、同q8のヒットを実Excelで確認。Issue #15の浸透結果無効化と入力変更も確認。
- 各試験は所有したfresh Excelで逐次実行。単一DLLだけのsandboxとmemory mapsの一致、終了後のExcelプロセス退出を確認。

## 本番規模4場の原式再監査

保存済みPython場から **primalの応力・速度と実メッシュだけ**を渡し、G1R19の元のVBA式で再組立・再監査しました。
材料・幾何・荷重・方針・場の完全roundtripも照合。求解0回、根区間・solver status・双対は取り込まず、既存の上下界と状態を変更していません。
下界のraw応力残差1e-7、元物理ゲート、下界λ>=1.000001／上界λ<=0.999999を保持しました。

| ケース | 側 | 要素数 | Fs | 元物理値λ | 元行列の等式 / cone残差 | 判定 |
|---|---|---:|---:|---:|---:|---|
| P06A | lower | 22,080 | 0.8468 | 1.00000150002666 | 1.687e-09 / 3.279e-09 | 合格 |
| P06A | upper | 12,753 | 0.85755 | 0.999997999931677 | 6.771e-08 / 5.073e-09 | 合格 |
| P06N | lower | 11,201 | 0.678279841819045 | 1.00004140787049 | 4.084e-10 / 0.000e+00 | 合格 |
| P06N | upper | 27,917 | 0.6876 | 0.999997999820129 | 2.204e-08 / 9.064e-10 | 合格 |

全行と全coneの独立replayも合格。監査時間はVBAの大規模組立を含み、ソルバー速度の比較には使いません。
[各場の原結果とSHA](https://github.com/tak063495-prog/excel-vba-rpfem/tree/main/versions/G1R19/validation/excel) を保存しています。
この確認は初期状態からの適応解析・Fs根探索による2%自動再現ではありません。

## 固定点の速度比較

同一PC・同一モデル・実451三角形・q8・Fs=1の上側探索について、従来経路とDLL先行を逐次各1回比較しました。
最終VBAソースと最終単一DLLによる測定です。全解析時間の短縮率や反復測定の統計とは区別します。

| ケース | 従来 | DLL先行 | 短縮率 | 採用した証拠 |
|---|---:|---:|---:|---|
| P06A | 29.56250秒 | 6.59375秒 | 77.7% | `original_dual_helper` |
| P06N | 29.91406秒 | 5.90625秒 | 80.3% | `original_primal` |

P06Aの原双対helperは探索符号の証拠で、追加の連続体上下界として扱いません。
P06Aの原LPのみの先行案は速度効果がなく、時間制限内のfresh原双対helperを加えた最終経路で改善しました。
その旧測定と、同runの無効GAP=2によるterminal fixture失敗は参考履歴として保持し、合格suiteには数えていません。
最終N固定点は別の最終ソースrunで測定・原全ベクトルの独立replayを完了しました。

## 配布ブック・証拠・検証範囲

- 62モジュールのUTF8・CP932内容とSOURCE_MANIFESTの正確なバイトSHAを照合。
- 配布ブックはcompile後に保存・再読込み・再compile。標準斜面・未解析、GAP=.02、DIRECT_DLL=0、precision-from-start=0。
- 「2%・速度優先の設定」はDIRECT_DLL=1、q8、根幅1e-4を設定し、材料・領域・荷重・解析方針と保存場のバイトを保持。
- 配布ZIPはDLL1個、全メンバのバイト・長さ・SHAをmanifestと照合。[WORKBOOK_VALIDATION](https://github.com/tak063495-prog/excel-vba-rpfem/blob/main/versions/G1R19/delivery/WORKBOOK_VALIDATION.json) / [ZIP_VALIDATION](https://github.com/tak063495-prog/excel-vba-rpfem/blob/main/versions/G1R19/delivery/ZIP_VALIDATION.json)。
- 生のログ・CSR・field・試験ブックはReleaseの `RPFEM_20261015_G1R19_Regression_Evidence.zip` に保存。確認用証拠ZIPにDLLは含めていません。
- [EXCEL_EVIDENCE_MANIFEST](https://github.com/tak063495-prog/excel-vba-rpfem/blob/main/versions/G1R19/validation/EXCEL_EVIDENCE_MANIFEST.json) は元runのSHAを保持し、参考履歴のFAILを合格と混同しません。入力runのバイトは未変更。

保存済みPythonの確定幅はP06A **1.2614779828%**、P06N **1.3647112865%**で2%基準に合格。
今回の追加精度実験から新しい確定上下界は採用していません。旧26ケース全部の本番計算、
初期状態からの本番2%再現、解析全体の速度短縮率は再確認していません。[残る範囲 #11](https://github.com/tak063495-prog/excel-vba-rpfem/issues/11)。
