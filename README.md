# excel-vba-rpfem

Excel VBAによる2次元RPFEM解析のソース、XLSM、検証資料です。

## G1R11 — P06N設定済み改訂版

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

## ライセンス

リポジトリの既存[LICENSE](LICENSE)を保持しています。
