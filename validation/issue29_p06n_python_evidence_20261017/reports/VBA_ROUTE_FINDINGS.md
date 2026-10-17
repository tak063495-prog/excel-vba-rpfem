# Issue #29：P06NのPython成功結果とFix4通常計算の差

確認日：2026-10-17（日本時間）。対象は `versions/G1R20_Fix4/src/vba` と、既存のPython実験・再現記録。今回の調査は読取りと静的抽出であり、Excel・最適化を実行していない。製品ソース・プリセット・DLLは変更していない。

## 現時点の結論

Pythonで報告したP06Nの対称GAP **1.9723805913%** は、下限5966要素・上限8739要素の**既知のメッシュ形状から物理場と根を新しく求めた結果**である。通常の粗い初期形状から、そのメッシュまで自動生成して2%以内を達成した証拠ではない。Fix4の速度プリセットの通常計算と同じ実験として扱うことはできない。

設定・幾何生成・数値経路の差はソースから確定できる。ただし、Issue #29の実行機からの完全なTriela生ログ・各側のメッシュ・全Fs試行結果は未入手であり、報告された未達・長時間の支配原因を一つに断定できない。改善プリセットへ変更するだけでP06Nが2%を達成することも未確認である。

## 設定の確定差

`RPX_Speed2pctPreset` は `RPX_AccuracyPreset` を呼び、`RPX_G20ImprovementPreset` はSpeedを呼んでから上書きする。下表はその明示代入の最終値であり、Excelブックの現在値を読み取ったものではない。ほかの既存設定は残る可能性がある。

| 設定 | 速度プリセット | G20改善プリセット | Fix4の根拠 |
|---|---:|---:|---|
| GAP | 0.02 | 0.02 | `RPX_UI.bas:491,518` |
| QUADRATURE / FS_TOL | 8 / 0.0001 | 8 / 0.0001 | `RPX_UI.bas:491` |
| ADAPT_CERT_QUADRATURE | 8 | 8 | `RPX_UI.bas:494` |
| DIRECT_DLL | 1 | 1 | `RPX_UI.bas:518` |
| ADAPT_PRECISION_FROM_START | 1 | 1 | `RPX_UI.bas:519` |
| ADAPT_BALANCED_LOWER_TARGET | 0 | 1000 | `RPX_UI.bas:521,547` |
| ADAPT_UPPER_FLOW_PASSES | 0 | 5 | `RPX_UI.bas:521,547` |
| ADAPT_LOWER_TARGET / ADAPT_LOWER_FAN_ROUNDS | 4000 / 1 | 0 / 0 | `RPX_UI.bas:494,548` |
| FS_TERMINAL_SCHEDULE | 0 | 1 | `RPX_UI.bas:521,549` |
| ACCURACY_DLL_UPPER_RECENTER | 0 | 1 | `RPX_UI.bas:522,550` |
| ACCURACY_DLL_DUAL_LINEAR_REPAIR | 0 | 1 | `RPX_UI.bas:523,551` |
| DIRECT_DLL_SECONDS | 5秒 | 30秒 | `RPX_UI.bas:524,554` |
| ACCURACY_DLL_DUAL_SECONDS | 20秒 | 60秒 | `RPX_UI.bas:524,554` |
| ACCURACY_DLL_SECONDS | 120秒 | 90秒 | `RPX_UI.bas:524,552` |
| ACCURACY_DLL_DUAL_REGULARIZATION | 0 | 1e-8 | `RPX_UI.bas:523,555` |

機械抽出した全明示代入と上書き元行は `PRESET_ASSIGNMENTS.json` に保存する。`extract_preset_assignments.py` は対象ソースの実バイト数・SHA256と、抽出器自身のSHA256を記録する。EnsureSettingRowなどの「欠落時だけ設定する」副作用は実行しておらず、Issue実行での実際の設定を証明するものではない。

## 幾何・数値経路の確定差

| 観点 | Pythonの成功記録 | Fix4ソースで確認した動作 | 意味・限界 |
|---|---|---|---|
| 初期条件 | `G1R20_RECIPE.md:14–29`：既知形状のxy/tri/material_idを読み、材料・荷重・境界は正規入力から再構築。旧物理場・旧Fsは数値seedにしない | `RPX_Adapt.bas:1729–1747`：TARGETの35%、256–600を初期目標として通常生成（定数は62–65） | 既知形状再求解とcold適応解析は別検証 |
| 成功下限の形状履歴 | `G1R20_RECIPE.md:29`：履歴付きnative1000→uniform4000→stress_fan5350→star5422→5594→5966 | `RPX_Adapt.bas:2238–2261`：G20 balancedは新しくTARGET1000を生成しred4。Speedの別候補は`1314–1366`：新しくTARGET4000生成＋angular fan最大1回 | いずれも成功した全履歴と同一ではない |
| 成功上限の形状履歴 | `G1R20_RECIPE.md:29`：2914形状へ5pass/115flip→local kinematic6106→directional8739。local/directionalの実thetaとより前の2914生成は未回収 | `RPX_Adapt.bas:2263–2302`：G20 flowは現在の上限場から辺flip。rn/re不変を要求（2291） | **成功形状を最初から全て生成する問題は未解決**。G20有効化だけで8739成功形状が再生されるとはいえない |
| LP/QP求解 | `python/accuracy_driver.py:258`はaffine縮約、`experiment_core.py:440–442`で物理場とdualを元空間に復元して監査 | `RPX_AccuracyNative.bas:138–169`はaffine=Falseの原問題を要求し、そのままDLLへ渡す | 原方程式の一致と数値経路・所要時間の一致は別。これ自体は物理式の不一致を示さない |
| DIRECT_DLLの意味 | Pythonは各点で縮約LP等を試し、監査済み符号を選択（`accuracy_driver.py:258–330`） | `RPX_Fs.bas:192–201`：DLL先行が監査合格なら即return、棄却ならHSD等へ。`RPX_AccuracyNative.bas:791–830`：棄却時に以前の状態を復元 | DIRECT_DLLは帯LUだけの高速化切替ではなく、Fs各点のソルバー選択順序を変える |
| c=0時のfallback | PythonはClarabel縮約問題＋同じ原式の監査 | `RPX_Fs.bas:186–190`、`RPX_Robust.bas:44–47`：c=0があればHSD。direct採用時はそのHSDを飛ばす | DLL先行の5秒候補の棄却・HSD所要時間が必要。DIRECT_DLL=1だけでHSDが一切走らないわけではない |
| q4/q8 | 成功再現の各物理場・根はq8 | 普通のEvaluateはprecision設定を尊重（`RPX_Adapt.bas:365–368`）。ただし上限pilotは`1180–1181`で無条件にq4/根許容0.001へ制限 | Issue実行がpilotを通ったかは実ログ要確認。最終q8証拠までq4を転用するという意味ではない |
| 元dualとQP dual | `G1R20_RECIPE.md:45–66`：QP dualは根の証拠にしない。最後は原dualの自由等式乗数だけ補正し、cone変数不変のまま原式再監査 | `RPX_AccuracyNative.bas:579–581,1004–1011,1132–1167`：補助dual破棄、同じ原dual補正と再監査。ただしSpeedは補正・recenter OFF | 速度設定では成功再現の重要な回復手順が無効。そこが実際のどのUNKNOWNを解決するかは別途試験が必要 |
| own-q8の再求解 | Pythonは同一メッシュ・同一点をcacheし、符号/物理場を区別（`accuracy_driver.py:225–258`） | `RPX_Adapt.bas:1457–1462`：terminal再利用は2%達成時だけ。未達時は`1588–1593`→`1532–1539`でcacheと根をresetし両側FsEndpointを新規実行 | 既にq8・自己根がある場合でも再求解し得る。長時間化の候補だが、Issueの時間内訳なしで支配要因とは断定不可 |
| 未達時の判定 | 原物理場・自己根・対称GAP条件が必要 | `RPX_Adapt.bas:1970–1977`：2%を超えたらSearchIncompleteとADAPT_GAP_TARGET_NOT_METを記録 | 2%の設定値は停止目標であり、任意の初期形状での達成保証ではない |

参照パスの省略：VBAは `versions/G1R20_Fix4/src/vba/`、Pythonは `validation/G1R19_python26_20261016/python/`、recipeは `validation/G1R19_python26_dual_linear_repair_20261016/G1R20_RECIPE.md`。

## 生ログがないため断定できない仮説

1. 速度プリセットで無効な候補・QP再中心化・dual補正が、今回のP06N未達の一因となった可能性。確定しているのは設定差であり、改善プリセットのcold成功ではない。
2. 粗い通常メッシュからの適応経路が成功形状のすべり面・応力表現へ到達せず、離散上下限幅が残った可能性。要素数だけでは形状・材料境界・メッシュIDの一致を証明できない。
3. direct5秒の監査棄却後のHSD、UNKNOWNの再試行、own-q8の再求解が時間を消費した可能性。ソルバー・stageごとの秒数と採用/棄却理由が必要。
4. pilotのq4経路の有無、通常q8経路の有無、実ブックの残存設定が探索を分岐させた可能性。関係する設定表・feature eligibility・adapt_evalが必要。

`versions/G1R20/README_JA.md:32`、`TEST_RESULTS.md:80` は既に通常初期状態からのVBA P06本番2%自動完走・全26例VBA再実行・解析全体の速度改善率を未確認と明記している。過去の固定点一致やPython保存場の監査を、今回のcold経路の成功へ読み替えない。

## 検証の優先順位

1. Issue機の完全な入力・設定、実行ソース/ブック/DLLのSHA、adapt履歴、上下各側のメッシュID/要素数/q/根、Fs全試行とnative/HSD時間を保存。最終表示GAPだけで判断しない。
2. 既存のPython証拠ZIPを展開し、**最適化0回の原q8再監査**で5966/8739の物理場・自己根・1.9723805913%を確認する。新規cold求解とは明記して区別する。
3. 同一入力・新しい出力フォルダーでSpeedとG20改善の経路を比較する。過去の物理場/Fsを暗黙にseedにせず、段階ごとの採用kind・形状・時間を比較する。
4. 同一の既知メッシュ・同一FsでPython/VBAの原行列と物理場を比較し、その後に各自己根まで比較する。既存の8739要素での一固定点成功は根・cold全体の検証ではない。
5. 判明した支配要因だけを変更候補とし、未達時のown-q8重複回避や補助問題の順序を検討する。同一メッシュ/モデル/q/方針/Fs/根証拠を厳密に確認し、原監査閾値・2%基準を緩めない。

今回の成果は原因調査と証拠公開に限定する。solver修正、プリセット変更、2%達成保証は含まない。
