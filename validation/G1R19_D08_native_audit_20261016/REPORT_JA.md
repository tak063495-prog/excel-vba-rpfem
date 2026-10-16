# D08 実Excel cold・最終原場の独立監査

`validation/G1R19_WaterCommon_excel_20261016/results/owned_excel_1792152100656896200` の最終 `00007_lower` / `00008_upper` を新しいPythonプロセスで読み、元のFs・元の場を無変換で原q8物理式へ再組立した。求解呼び出しは0回。原基準の上下界・荷重・境界・材料・policy・meshと根metadataを確認した。

| 項目 | 下界 | 上界 |
|---|---:|---:|
| Fs | 1.3015815066522918 | 1.3260274687490892 |
| 要素 | 429 | 429 |
| q | 8 | 8 |
| 独立物理場 | PASS | PASS |
| 根の絶対幅 | 0.0001269216486255 | 0.0001309656759259 |
| 根の正規化幅 | 0.0000975086538932 | 0.0000987703096449 |

対称GAPは **1.8607001517844293%**。根の既存契約は `(b-a)/max(1,(a+b)/2)<=1e-4` で、両側とも満たす。絶対幅1e-4の契約とは表現しない。実Excel coldは883.9523024秒で900秒内に正常完了、owned PID9228は退出済み。過去のPython場・根はimportされていない。

下界のraw equilibriumは1.78523e-8、internal tractionは2.98774e-9、yieldは0、荷重倍率は1.0000077069854048。上界のflow/jump違反は0、仕事誤差は1.08425e-10、原散逸値は0.9999890760382877。原物理門を変更せずに合格した。

最終ブックの浸透結果557節点をXMLから読み、全保存水頭が正確に3、保存された符号付き水圧が9.81(3−y)であることを検証した。解析のsuction=0に従い、共通体積力は水位以下で−γ+9.81、水位以上で−γ、固定荷重成分は0となる。水位を横断する三角形はなく、面積158、合力[0,−2571.4]を両meshで確認した。乾燥モデルへのfallbackはない。

原full-model文字列の節点・要素・材料・体積力・全辺と、保存binおよび独立再構成topologyが一致した。外周の幾何constraint57と内部水位の幾何constraint38を分離した。内部水位保護辺は物理的な固定境界へ変換せず、外周物理BCと荷重は正本D08のまま照合した。

正本input、凍結62ソース、ハーネス、DLL、WaterCommon manifest（e25c0748…）、最終ブック・両場・両meta等74ファイルは監査前後でbyte不変。最終ブックは `summary.files` のSHAへ対応する。`configured_workbook_sha256` は解析前の別checkpointであり、解析出力の保存後SHAと混同しない。

この監査は採用端点の全原物理場とown-rootの同一性・幅・完了を独立に確認した。棄却側全端点の原dual/rayを再求解・再証明したとは主張しない。59個のliteral CSR再生は元ハーネスの結果として区別し、本checkerがもう一度59件全部を監査したとの主張はしていない。

結果: `COLD_NATIVE_D08_INDEPENDENT_AUDIT.json`。再現と求解なしの小テスト:

```powershell
D:/Work/rpfem-venv/Scripts/python.exe -X utf8 validation/G1R19_D08_native_audit_20261016/audit_cold_water_native.py
D:/Work/rpfem-venv/Scripts/python.exe -X utf8 -m unittest discover -s validation/G1R19_D08_native_audit_20261016 -p test_water_auditor.py -v
```
