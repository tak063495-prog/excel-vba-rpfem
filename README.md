# excel-vba-rpfem

Excel VBAによる2次元RPFEM解析のソース、XLSM、検証資料です。

剛塑性有限要素法による**極限解析の下界・上界**を求め、強度低減安全率Fsまたは極限荷重倍率を出力します。変位・沈下量を求める弾塑性解析ではありません。非関連流れ則の入力ではDavis等価関連モデルを使うため、結果の解析政策も確認してください。

## 使い方・原理・注意事項

|資料|内容|
|---|---|
|[使い方](docs/USAGE_JA.md)|動作環境、ダウンロード、P06Nの実行、各入力表・設定・出力の読み方、支持力モード、ログの保存|
|[解析の原理](docs/PRINCIPLES_JA.md)|剛塑性と上下界、2次三角形要素、Mohr–Coulomb条件、強度低減・Davis変換、SOCP・内点法・HSD、適応メッシュとコードの対応|
|[注意事項・不具合時の確認](docs/NOTES_JA.md)|検証の範囲、GAPと根幅、解析領域・多層・剛体界面、c=0/小さいc、数値失敗、速度図・応力係数、マクロ・復元|

### 最初の実行

1. 下記のD03NまたはD07A設定済みXLSMをダウンロードし、ローカルの作業フォルダーへ別名で保存します。
2. 入手元・ハッシュを確認し、組織の規則に従ってそのファイルのマクロを実行できる状態にします。通常の解析にPythonやVBAプロジェクトアクセスの許可は不要です。
3. 「実行案内」「材料データ」「設定」を確認し、「要素定義」の**「要素作成②」→「図を表示」→「解析」**を使います。「解析」は入力からメッシュを再生成するため、要素作成は形状の事前確認用です。
4. 「解析結果」の下界・上界・区間幅・解析政策と、ログ末尾の判定を確認します。`PARTIAL_CERTIFIED`は目標未達を含みます。詳しい操作は[使い方](docs/USAGE_JA.md)を参照してください。

**最新版G1R15では、G1R14（定常浸透・間隙水圧の一方向連成）に、浸透入力シートの整理・地下水面からの水圧入力・水平震度法（地震）を追加した（実Excelで全回帰済み：連成OFF・kh=0でD03N・D07AのFsがG1R14と全桁一致）。G1R14では連成OFFでD03N・D07AのFsがG1R13と全桁一致した。G1R13の修復・制限は[G1R13報告](versions/G1R13/REPORT.md)、G1R12時点の25ケース結果（23/25取得・11/25目標達成）は[過去25ケース結果](validation/G1R12_native25_20261009/REPORT.md)を参照し、区別する。図の矢印は規準化された速度機構であり、変位量ではない。**

## G1R15 DLL2 — 64bit Excel用の高速化版

2026年10月12日。G1R15のLDL・DD数値核と並べ替えをDLLへ移植した版です。堤防例題で **622.12秒（元VBAの既存記録）→443.89秒（前回DLL）→400.16秒（DLL2）**。元VBA比 **35.68%短縮・約1.55倍**、前回DLLから **9.85%短縮**でした。同じPCの既存記録との比較で、元VBAは再計測していません。

- **[XLSM・DLL・ソース・検証結果の一式ZIPをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R15_DLL2_Levee_Checked.zip)**
- [最新DLL（rpfem_native_x64.dll）をダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R15_DLL2/delivery/rpfem_native_x64.dll) / [対応XLSMをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R15_DLL2/delivery/RPFEM_20261012_G1R15_DLL2_Levee_Checked.xlsm)
- [使い方](versions/G1R15_DLL2/README.md) / [検証報告](versions/G1R15_DLL2/REPORT.md) / [導入・復旧](versions/G1R15_DLL2/IMPORT_ROLLBACK_JA.md) / [VBAソース](versions/G1R15_DLL2/src/vba/) / [DLLのC・C++・ASMソースとライセンス](versions/G1R15_DLL2/native/) / [差分](versions/G1R15_DLL2/changes.diff) / [SHA-256](downloads/SHA256SUMS.txt)

ZIPを展開し、`delivery/` のXLSMとDLLを同じフォルダーに置いてください。**DLLだけを元G1R15へ追加しても高速化経路には入りません。対応XLSMを使うか、導入手順でVBAを更新してください。** 前回DLL版を使っていたExcelは終了してから新しい版を開きます。DLL不在・読込不可・版不一致ではVBAへ自動切替します。

64bit Excelで検証済み。構成則・Davis変換・Fs定義・許容差・監査を維持し、堤防の53求解・603監査は元VBAと前回DLLのログ記録精度で一致しました。Fs下界1.224448、上界1.234308、上下界差0.802056%。今回の実解析は堤防1例で、DLL2の25ケース全回帰と、DDを多用する実解析の高速化率は未確認です。

## G1R15 — 浸透入力の整理・地下水面入力・水平震度法（最新版）

2026年10月11日。詳細は[G1R15 README](versions/G1R15/README.md)・[検証報告](versions/G1R15/REPORT.md)。

- **地震（水平震度法）**：`設定`シート B15＝水平震度 kh（0＝なし、目安0.14〜0.25）、B16＝慣性力の向き 右(+x)／左(-x)。浸透とは独立で、浸透なしでも使える。全要素に ±kh·γ を水平に作用。
- **浸透入力**：`浸透入力`B3で「水位・水頭の条件から計算する／地下水面の位置を入力する（斜面に沿った流れ cos²θ）／地下水面の位置を入力する（静水圧）／地下水面に合うよう境界の水位を推定する」を選ぶ。H3に必要な入力を表示。連成スイッチはD4（G1R14のN19から移動、自動移行）。
- 全回帰（kh=0・連成OFF）：D07A 1.26514209147729／1.27757720458799（1511秒）、D03N 1.11795089664424／1.20943742542898（9415秒）＝G1R14と全桁一致。
- 最適化の途中段で下界の探索区間が見つからない場合も解析を止めず、直前の検証済みの上下界を保持して続行（下界探索は前段下界の15%で打切り。根拠は[検証報告](versions/G1R15/REPORT.md)）。
- スモーク試験（D07A・浸透なし・200要素）：kh=0 1.2223／1.3640、kh=0.15谷側 0.9790／1.0986。堤防例題・連成ON：kh=0 1.6977／1.7972、kh=0.15 1.2172／1.2695。堤防例題・連成ON＋kh=0.15右＋最適化ON：1.2244／1.2343（差0.80%、596要素、622秒）。
- [ソース・差分・検証資料ZIPをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R15_Seismic_20261011.zip) / [配布ファイルのSHA-256](downloads/SHA256SUMS.txt)
- [D07A](versions/G1R15/delivery/RPFEM_20261011_G1R15_D07A_Seismic.xlsm) / [D03N](versions/G1R15/delivery/RPFEM_20261011_G1R15_D03N_Seismic.xlsm) / [堤防例題](versions/G1R15/examples/RPFEM_20261011_G1R15_Levee_Example.xlsm) / [VBAソース](versions/G1R15/src/vba/) / [差分](versions/G1R15/changes.diff) / [導入・復旧](versions/G1R15/IMPORT_ROLLBACK_JA.md)

## G1R14 — 浸透流解析・間隙水圧の一方向連成（前版）

2026年10月11日。三角形メッシュの定常浸透流解析（不圧・自由水面）を追加し、その間隙水圧を有効応力として下界・上界のFs探索へ反映できるようにした。図シートに自由水面と推定すべり面（速度ジャンプの連続経路と当てはめ円弧）を重ねて描く。既存の構成則・荷重・数値核・監査は変更していない。変更は既存2モジュールへ各1行と新規3モジュール（`RPX_Seep`・`RPX_SeepUI`・`RPX_Water`）。

### 使い方

1. 下記のD03NまたはD07A浸透対応XLSMを別名で保存して開く。
2. 「浸透入力」シートで材料ごとの透水係数（A:C）と水理境界（E:K。要素定義の領域・点で範囲を指定、種別は水頭／水位／水位+浸出面／浸出面／流量／不透水）を入力し、「浸透解析」を押す。
3. 「図を表示」で自由水面を確認する。連成する場合は N19 を 1 にしてから、通常どおり「要素定義」の「解析」を実行する（γは湿潤・飽和の全単位体積重量）。
4. N19=0（既定）ではG1R13と同じ解析になる。詳しくは[手順書](versions/G1R14/docs/浸透連成_手順書.md)。

### 検証（実Excel）

|試験|結果|
|---|---|
|D07A 連成OFF|下界1.26514209147729・上界1.27757720458799（1513秒）＝G1R13公開値|
|D03N 連成OFF|下界1.11795089664424・上界1.20943742542898（9381秒）＝G1R13公開値|
|矩形ダム（H1=10, H2=2, L=20）|52反復、Q=2.4181（Dupuit 2.4000、+0.75%）|
|全没水：γ=20＋連成 と γ'=10.19|Fs 2.065489／2.095742で一致（G1R12系列で実施）|
|堤防例題（連成）|浸透29反復・流入＝流出、Fs 下界1.7104・上界1.7189、すべり円 中心(25.16,17.29) R=10.84|
|重ね描き自己テスト|合格（連続すべり面39節点、円 (15.77,15.32) R=10.36／真値(16,16) R=11）|

### 制限

- 定常・一方向連成。過剰間隙水圧・圧密・非定常は対象外。浸出点の高さはkr平滑化でやや低めに出る。
- 推定すべり面は上界速度場の可視化で、Fsの根拠ではない。
- 水面下の剛体要素には水圧を作用させない。
- D03Nの全解析は約2.6時間。複数のExcelを同時に動かすとCOM接続が切れた例がある。

### ダウンロード

- [D03N浸透対応XLSMをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R14/delivery/RPFEM_20261011_G1R14_D03N_Seepage.xlsm)
- [D07A浸透対応XLSMをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R14/delivery/RPFEM_20261011_G1R14_D07A_Seepage.xlsm)
- [堤防例題XLSMをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R14/examples/RPFEM_20261011_G1R14_Levee_Example.xlsm)
- [ソース・差分・検証資料ZIPをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R14_Seepage_20261011.zip)
- [検証報告](versions/G1R14/REPORT.md) / [VBAソース](versions/G1R14/src/vba/) / [統合差分](versions/G1R14/changes.diff)
- [インポート・復旧手順](versions/G1R14/IMPORT_ROLLBACK_JA.md) / [要求と実装の対応](versions/G1R14/ISSUE_TO_FIX_JA.md)

## G1R13 — D03N・D07A修復版（前々版）

2026年10月10日。D03Nの参照仕事の桁落ちと候補の正規化、D07Aの荷重段差節点の2三角形fanを修復した。構成則、荷重、Davis変換、数値核・監査許容値、cold-start政策を維持した。

|ケース|下界Fs|上界Fs|GAP|精度区分|
|---|---:|---:|---:|---|
|D03N（Davis等価）|1.11795089664|1.20943742543|7.861733%|FS_AVAILABLE_PARTIAL|
|D07A（関連流れ則）|1.26514209148|1.27757720459|0.978096%|FS_AVAILABLE_PARTIAL|

**両ケースとも実ExcelでFsを取得し、保存した上下界場を独立Pythonで監査した部分結果。D03NはGAP目標1%と既存の根有効性条件、D07Aは採用上界の根幅が未達である。精度目標達成とは扱わない。**

両XLSMに同じ修正版VBAが入っており、どちらもD03N・D07Aの両ケースに対応する。違いは入力済みの形状・材料・境界条件・荷重・解析方針と保存結果。ケースを切り替える場合はこれらを一式変更し、通常の「解析」から再実行する。D03NはDAVIS_EQUIVALENT、D07AはASSOCIATED。

- [D03N設定済みXLSMをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R13/delivery/RPFEM_20261010_G1R13_D03N_Verified.xlsm)
- [D07A設定済みXLSMをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R13/delivery/RPFEM_20261010_G1R13_D07A_Verified.xlsm)
- [ソース・差分・検証資料ZIPをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R13_D03N_D07A_Checked.zip)
- [修復・検証報告](versions/G1R13/REPORT.md) / [VBAソース](versions/G1R13/src/vba/) / [統合差分](versions/G1R13/changes.diff)
- [インポート・復旧手順](versions/G1R13/IMPORT_ROLLBACK_JA.md) / [問題と修正の対応](versions/G1R13/ISSUE_TO_FIX_JA.md)
- [試験台帳](versions/G1R13/results/test_ledger.csv) / [保存場の独立監査](versions/G1R13/results/portable_audit.json)
- [配布ファイルのSHA-256](downloads/SHA256SUMS.txt)

今回G1R13で全解析したのはD03N・D07Aの2ケース。25ケースの粗メッシュ比較ではD07A以外24ケースのモデルが8バイトDouble値まで一致した。仕事修復15条件、荷重段差11条件、実Fs経路7失敗注入、既存12保存場の再監査を実施した。25ケース全部のG1R13全解析回帰や全解析の速度向上率は確認していない。旧G1R12の23/25という結果と今回の2ケースを、同一版25/25の合格へ合算しない。

## G1R12 — Issues #1 / #2 修正版（前版・保全）

2026年10月9日。2件の指摘を実コードとExcelの反例で確認し、修正しました。適応解析の最良上界メッシュと速度場を出力前に一致させ、`NEWTON_FACTORIZATION_FAILED pivot=…`の実エラー形式で同一問題のHSD救済を1回行います。構成則・物理監査・数値カーネル・許容誤差・G1R11の高速化政策を維持しています。

- [P06N設定済みXLSMをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R12/RPFEM_20261009_G1R12_P06N_IssueFix.xlsm)
- [ソース・差分・検証資料ZIPをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R12_IssueFix_20261009.zip)
- [修正・検証報告](versions/G1R12/REPORT.md) / [VBAソース](versions/G1R12/src/)
- [手動反映・G1R11へ戻す手順](versions/G1R12/IMPORT_ROLLBACK.md)
- [試験手順](versions/G1R12/TEST_PROCEDURES.md) / [試験結果](versions/G1R12/TEST_RESULTS.csv) / [差分](versions/G1R12/changes.patch)
- [Issue対応表](versions/G1R12/ISSUE_FIX_MAP.csv) / [25ケースの実行結果](validation/G1R12_native25_20261009/RESULTS.csv) / [Issue修正時点の台帳（保全）](versions/G1R12/excel_acceptance_g1r12.csv)
- [配布ファイルのSHA-256](downloads/SHA256SUMS.txt)

旧版の601/1000要素の配列エラー、同数・別座標の誤対応、詳細エラー時の救済スキップを実Excelで再現しました。修正版の出力・停止条件・実HSD救済を検証し、小規模C=5・φ=ψ=45のFs根と通常場は旧版と完全一致しました。保存場7件をPythonで独立監査し、小規模の支持力通常・適応経路、完成XLSMの保存後コンパイル・全46モジュール照合も実施しました。人工的な所有者・失敗注入の試験は物理求解と区別しています。

### 2026年10月9日の25ケース実行

間隙水圧を除く25ケースすべてを実Excelで実行しました。監査済みFs区間は23/25、GAP 1%と根探索目標の達成は11/25です。最終Fs未取得はD03N・D07Aです。 取得Fsのうち12ケースは精度目標が未達です。全支持力・剛体回帰は今回の対象外です。

- [25ケースの原ログ・監査・再現コードZIP](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R12_Native25_Evidence_20261009.zip)
- [結果XLSM ZIP①（D01A～P02N、12件）](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R12_Native25_Workbooks_01_20261009.zip)
- [結果XLSM ZIP②（P03A～P10A、13件）](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R12_Native25_Workbooks_02_20261009.zip)

XLSMは2つの独立したZIPです。それぞれ展開して使用でき、結合は不要です。ブック内容は検証済みの一括ZIPと同一です。
- [報告書](validation/G1R12_native25_20261009/REPORT.md) / [結果CSV](validation/G1R12_native25_20261009/RESULTS.csv)

製品VBAは同じG1R12で、試験用証拠モジュールを追加したコピーを使いました。D07Aの荷重端点メッシュ、D03NのDD積underflowについてPythonで原因・対策を検証しましたが、修復のVBA実装はまだ行っていません。G1R11/G1R12の既存公開XLSM・ZIPは保全しています。

## G1R11 — 前版（保全）

2026年10月8日に登録した版です。従来の再配置・細分化を先に行い、その後の追加細分化を最大1000要素・1回に制限します。追加段階の多層モデルでは監査済み上界を下界Fs探索の試行位置の目安に使い、符号を実求解します。構成則・Davis変換・監査・許容誤差・G1R9のDD高速化を維持しています。

- [P06N設定済みXLSMをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/versions/G1R11/RPFEM_20261007_G1R11_P06N_Staged_1000.xlsm)
- [ソース・差分・検証資料ZIPをダウンロード](https://github.com/tak063495-prog/excel-vba-rpfem/raw/refs/heads/main/downloads/RPFEM_G1R11_Staged_20261007.zip)
- [実装・検証報告](versions/G1R11/REPORT.md)
- [VBAソース](versions/G1R11/src/)
- [手動反映・復元手順](versions/G1R11/IMPORT_ROLLBACK.md)
- [試験手順](versions/G1R11/TEST_PROCEDURES.md) / [試験結果](versions/G1R11/TEST_RESULTS.csv)
- [変更差分](versions/G1R11/changes.patch) / [問題と修正の対応表](versions/G1R11/ISSUE_FIX_MAP.csv)
- [配布ファイルのSHA-256](downloads/SHA256SUMS.txt)

### 検証範囲

Excelの保存後再読込み・コンパイル・コード照合、P06Nの1000要素における下界Fs=0.655と上界Fs=0.77の固定点求解、独立した原式監査、小規模単層C=5・φ=ψ=45のFs根回帰、失敗注入による制御試験を実施しました。制御試験の代替値・代替監査は物理求解とは区別しています。

**全P06N適応解析の時間短縮率、GAP 1%到達、間隙水圧を除く25ケースの全回帰は未確認です。** 個別台帳は[excel_acceptance_g1r11.csv](versions/G1R11/excel_acceptance_g1r11.csv)にNOT_RUNとして残しています。固定点試験を全Fs解析の合格として扱っていません。

### 文字コードと再現

`src/`はUTF-8編集用です。VBE手動インポートにはCP932・CRLFの`import_cp932/`を使います。`rollback_cp932/`には改良元G1R10へ戻す3モジュールを収録しています。ThisWorkbookやSheetを新規標準モジュールとしてインポートする変更はありません。

`versions/G1R11/`には元の納品ZIPの内容をそのまま展開しています。Python・Excel検証のスクリプト、実行結果、保存場を含みます。再実行時に必要な保全原本と環境の参照先は[手順書](versions/G1R11/IMPORT_ROLLBACK.md)を参照してください。GitHubへ登録する際にFEM試験を再実行したものではありません。

使い方・原理・注意事項は2026年10月9日に追加しました。これらの説明はリポジトリの`docs/`にあります。2026年10月7日付の配布XLSM・ZIPと、収録済みの検証結果はそのまま保持しています。

## ライセンス

リポジトリの既存[LICENSE](LICENSE)を保持しています。
