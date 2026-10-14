# バージョンと配布物

最新の改善版は **[G1R18 Release](https://github.com/tak063495-prog/excel-vba-rpfem/releases/tag/G1R18)** です。配布ZIPにはXLSM、既存DLL、追加の精度回復DLL、ソースと導入手順が含まれます。

|版・記録|位置付け|案内|
|---|---|---|
|G1R18・2026-10-14|現在の精度改善版。P06の1%目標は継続課題|[配布・設定・検証範囲](G1R18/README_JA.md) / [ZIP](https://github.com/tak063495-prog/excel-vba-rpfem/releases/download/G1R18/RPFEM_20261014_G1R18.zip) / [SHA-256](G1R18/SHA256SUMS.txt)|
|G1R18 Issue #12・2026-10-14|下界キャッシュの証拠メタデータ修正（`LOWER_MODEL_CACHE=1` 時）|[修正と検証](G1R18_Issue12/README.md) / [ZIP](../downloads/RPFEM_G1R18_Issue12_20261014.zip)|
|G1R17|前版。P06粗メッシュ修復と実Excelでの26ケース実測|[版の説明](G1R17/README.md) / [26ケース実測](../docs/benchmarks/G1R17_native26_20261013/README.md)|
|G1R17 Issue #8|水理収束判定の修正。G1R18へ取り込み済み|[水理修正と検証](G1R17_Issue8/README.md)|
|G1R16 DLL5|過去のDLL高速化・回帰確認|[配布と検証](G1R16_DLL5/README.md)|
|G1R13|D03N・D07Aの修復時点の保存版|[配布と結果](G1R13/README.md)|
|G1R12|25ケースを実行した当時の保存版|[報告](G1R12/REPORT.md)|

G1R18では旧表で精度目標未達だった16ケース中14ケースのPython独立再監査で目標達成を確認しました。P06Aの幅は8.612876%から1.261478%、P06Nは12.038014%から1.364711%へ縮小しています。比較元は[G1R17の実Excel保存場](../docs/benchmarks/G1R17_native26_20261013/data/RESULTS.csv)、最新値は[Python実験のq8再監査](../validation/accuracy_20261013/snapshots/20261014_progress7_audit/REPORT_JA.md)です。

VBA/DLLは小規模数値回帰と配布ブックの確認を完了しています。P06の本番規模での再現と1%達成は[Issue #11](https://github.com/tak063495-prog/excel-vba-rpfem/issues/11)で継続します。

`G1R18_RayGate`はG1R18へ統合した実験候補の保存記録です。通常の導入にはG1R18の配布ZIPを使用してください。他の過去版も各ディレクトリに保持しています。
