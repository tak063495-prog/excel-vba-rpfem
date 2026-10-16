# G1R19 2%候補 — D01A/D02Aの停滞修正

公開済みG1R19を上書きしない4モジュールの候補です。基盤はG1R19とmainのIssue #11修正です。DLLは公開G1R19の単一DLL（SHA256 `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`）を使用し、ABI・材料・荷重・Davis方針・原物理門は変更していません。

修正前のD01Aでは、全c=0の散逸指標がゼロになり、約0.577〜0.700の弱い上下界を繰り返していました。対応するc0斜面用の既存上界候補は無効で、呼出位置も高価な適応処理の後でした。またq4前段ではDLL回復が使われず、q8の下界でも有限の荷重rayを物理候補として再監査できない経路がありました。後段のDLL実使用はログで確認しており、「DLLが全く使われなかった」という説明ではありません。

変更は以下です。

- `RPX_Adapt`：元の散逸指標の総量がゼロのときだけ既存のひずみ・速度jump指標を使います。正の元指標はそのままです。
- `RPX_Adapt` / `RPX_UI`：適格な乾燥・重力のみ・単一c0材料・6頂点斜面で、既存8三角形の上界候補を適応処理の早期に検証します。深さは斜面高さの0.001倍です。解析的初期値から上下界を決めず、元q8場とそのメッシュ上の独立根で採用します。下界メッシュは別に保持します。
- `RPX_AccuracyNative` / `RPX_Fs`：元の自由荷重下界問題のcostが厳密に−λ、b/hの全成分が厳密に0の場合だけ、有限の同Fs primalコピーを正倍率でλ=1.000003へ調整します。元PrimalAuditとFieldGateを通ったコピーだけを物理証拠として使い、solver statusは採用根拠にしません。不合格は従来解法へ戻り、元の場・境界・dualを保持します。
- `RPX_Fs`：符号だけを保証する証拠端点は、任意倍率の目的値をsecant補間に使わず二分します。元の根の採用関数・許容幅は変更していません。

通常設定の `ADAPT_PRECISION_FROM_START` は0です。新しい「2%・速度優先の設定」を選ぶと、検証済み条件 `GAP=.02 / QUADRATURE=8 / DIRECT_DLL=1 / ADAPT_C0_WEDGE_AUTO=1 / ADAPT_PRECISION_FROM_START=1` になります。積分・根幅の粗段制限を最初から外し、q8の原監査を使います。通常設定の既定値とプリセット適用後の値は別です。

検証用ブックでは「設定」シートの「2%・速度優先の設定」を選び、元の材料・外形・境界と解析方針を確認してから実行します。Pythonや過去場を読み込む操作はありません。下記は既存場・根・statusを一切輸入しない新規Excel計算の結果です。

| ケース | 実計算時間 | 下界 / 上界 | 対称GAP | 下界 / 上界要素数 |
|---|---:|---|---:|---:|
| D01A | 358.120秒 | 0.57733154296875 / 0.5774156487217974 | 0.014567% | 369 / 8 |
| D02A | 168.791秒 | 1.5 / 1.500051451551501 | 0.003430% | 360 / 8 |

両側q8、両側own-root有効、原FieldGate合格、search-incomplete=falseです。根幅は既存の `FS_TOL × max(1, bracket midpoint)` 契約です。D01Aの幅は下6.1035e−5/上6.2505e−5、D02Aは下1.2207e−4/上1.2500e−4で、D02Aを絶対幅1e−4以下とはしていません。専用Excelはすべて終了しています。

- [実計算の指標とソースSHA](COLD_REGRESSION_20261016.json)
- [D01A実Excelの全記録](results/owned_excel_1792150058120949800/summary.json)、[D02A実Excelの全記録](results/owned_excel_1792149830248919100/summary.json)
- [D01Aの同一メッシュ・同一Fsの独立Python監査](../../validation/G1R19_D02_route_audit_20261016/COLD_NATIVE_D01_INDEPENDENT_AUDIT.json)、[D02A独立監査](../../validation/G1R19_D02_route_audit_20261016/COLD_NATIVE_D02_INDEPENDENT_AUDIT.json)

独立監査は原材料・原荷重・原境界から同じq8問題を組み直し、採用場を再確認しています。native full original CSRはD01A 37件、D02A 36件を再生し全合格です。これは全26ケースのVBA完走を示すものではありません。全26ケースのPython 2%検証は進行中です。水理共有メッシュ、追加bootstrap、compile cacheの候補はこの4モジュールへ統合していません。

再生成・試験：

```powershell
python versions/G1R19_2pct/tools/current_sources.py --output-root <新しいフォルダー>
python versions/G1R19_2pct/tests/test_overlay_sources.py
python versions/G1R19_2pct/tests/run_adaptive_excel.py --case D02A --guards-only --ray-unit
python versions/G1R19_2pct/tests/run_adaptive_excel.py --case D01A --precision-from-start --budget 900
python versions/G1R19_2pct/tests/run_mesh_export_excel.py --preset-only
```

Excel試験は新規所有プロセス1本だけを使い、既存Excelがあれば起動を拒否します。900秒は協調停止要求の時刻で、単一分解が戻るまでを含む厳密なwall上限ではありません。中止した旧q4比較は未達のまま別ディレクトリに保存し、成功結果へ混ぜていません。

P06A/Nについては求解せず、新規GenerateMeshからTARGET350と1000を出力しました。双方の粗段は400要素、fullは1078要素でした。[mesh-only記録](results/mesh_only_1792150651304041200/summary.json)には原入力・材料・荷重・境界・単位とSHAがあり、場・Fs・rootは含みません。PythonとVBAの開始メッシュ比較用であり、P06の2%達成証拠ではありません。
