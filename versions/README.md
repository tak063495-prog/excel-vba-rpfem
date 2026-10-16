# バージョンと配布物

最新の改善版は **[G1R20 Release](https://github.com/tak063495-prog/excel-vba-rpfem/releases/tag/G1R20)** です。配布ZIPにはXLSM、統合DLL1個、ソース・検証範囲・導入手順が含まれます。

|版・記録|位置付け|案内|
|---|---|---|
|G1R20・2026-10-16|P06改善経路のVBA実装。Python新規場はA1.8174%/N1.9724%、VBAは途中経路比較まで|[導入・確認範囲](G1R20/README_JA.md) / [検証結果](G1R20/TEST_RESULTS.md) / [Pythonコード](G1R20/PYTHON_VALIDATION.md) / [ZIP](https://github.com/tak063495-prog/excel-vba-rpfem/releases/download/G1R20/RPFEM_20261016_G1R20.zip)|
|G1R19・2026-10-15|2%基準の高速化・単一DLL版|[配布・設定](G1R19/README_JA.md) / [回帰・速度](G1R19/TEST_RESULTS.md) / [ZIP](https://github.com/tak063495-prog/excel-vba-rpfem/releases/download/G1R19/RPFEM_20261015_G1R19.zip) / [SHA-256](G1R19/SHA256SUMS.txt)|
|G1R18・2026-10-14|前版の精度改善。1%判定は当時の記録|[配布・設定・検証範囲](G1R18/README_JA.md) / [ZIP](https://github.com/tak063495-prog/excel-vba-rpfem/releases/download/G1R18/RPFEM_20261014_G1R18.zip) / [SHA-256](G1R18/SHA256SUMS.txt)|
|G1R18 Issue #12・2026-10-14|下界キャッシュの証拠メタデータ修正（`LOWER_MODEL_CACHE=1` 時）|[修正と検証](G1R18_Issue12/README.md) / [ZIP](../downloads/RPFEM_G1R18_Issue12_20261014.zip)|
|G1R17|前版。P06粗メッシュ修復と実Excelでの26ケース実測|[版の説明](G1R17/README.md) / [26ケース実測](../docs/benchmarks/G1R17_native26_20261013/README.md)|
|G1R17 Issue #8|水理収束判定の修正。G1R18へ取り込み済み|[水理修正と検証](G1R17_Issue8/README.md)|
|G1R16 DLL5|過去のDLL高速化・回帰確認|[配布と検証](G1R16_DLL5/README.md)|
|G1R13|D03N・D07Aの修復時点の保存版|[配布と結果](G1R13/README.md)|
|G1R12|25ケースを実行した当時の保存版|[報告](G1R12/REPORT.md)|

G1R18では旧表で精度目標未達だった16ケース中14ケースのPython独立再監査で目標達成を確認しました。P06Aの幅は8.612876%から1.261478%、P06Nは12.038014%から1.364711%へ縮小しています。比較元は[G1R17の実Excel保存場](../docs/benchmarks/G1R17_native26_20261013/data/RESULTS.csv)、最新値は[Python実験のq8再監査](../validation/accuracy_20261013/snapshots/20261014_progress7_audit/REPORT_JA.md)です。

2026-10-15からは2%を合格基準とし、1%を狙う追加実験は停止しました。保存済みPython場のP06A/P06Nは新基準に合格します。VBA/DLLの回帰と、本番計算の初期状態からの精度再現は区別し、[Issue #11](https://github.com/tak063495-prog/excel-vba-rpfem/issues/11)で管理します。

`G1R18_RayGate`はG1R18へ統合した実験候補の保存記録です。G1R20には新しい改善候補経路を含み、確認範囲を導入手順に明記しています。他の過去版も各ディレクトリに保持しています。
