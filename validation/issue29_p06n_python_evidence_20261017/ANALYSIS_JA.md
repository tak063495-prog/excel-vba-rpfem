# Issue #29：P06Nの2%証拠と通常解析の未達をどう読むか

確認日：2026-10-17（日本時間）。製品の起点はG1R20 Fix4、参照コードはmain `1ab383bbc1764249a4320ca28ed3ac6682c8fbac`です。

## 比較している実験

|項目|過去のPython成功対|Issue #29のFix4実測報告|
|---|---|---|
|下界／上界|0.6767771508826014／0.6902587260388777|約0.621227／0.844666|
|対称GAP|1.9723805913%|約30.5%|
|最終要素数|下5966／上8739|下1290／上1290|
|出発点|過去の適応履歴で得た形状だけを読み、新しい物理場・自己根を求解|400要素から通常の適応細分化|
|経路|複数の原LP・QP・原dual・線形補正componentを後で束ねた対|`RPX_Speed2pctPreset`、`DIRECT_DLL=1`|
|実行時間の扱い|単一cold driverの正常完走時間ではない|16193秒との報告|

右列の詳細は、ユーザーが指定した[調査コメント](https://github.com/tak063495-prog/excel-vba-rpfem/issues/29#issuecomment-6233406927)に基づきます。今回のTrielaのstatus・analysis・stage原ファイルを直接再集計した値ではありません。左列は当時のバイトを保持した入力・メッシュ・場を今回独立再監査しています。両者の区別は `reports/ISSUE29_CAPTURE_PROVENANCE.json` に記録しました。

## 1. 速度プリセットは改善候補の一部を無効にする

これはソースで確認できる設定差です。G20改善プリセットは速度プリセットを呼んだ後、次を上書きします。

|設定|Speed2pct|G20Improvement|
|---|---:|---:|
|GAP|0.02|0.02|
|DIRECT_DLL / 最初から高精度|1 / 1|1 / 1|
|balanced下界のTARGET|0|1000|
|上界flowのpasses|0|5|
|終端の探索順制御|0|1|
|上界QP再中心化|0|1|
|原dualの線形補正|0|1|
|直接DLLの1回予算|5秒|30秒|
|原dual復元の1回予算|20秒|60秒|

根拠：[RPX_UI.bas L511以降](https://github.com/tak063495-prog/excel-vba-rpfem/blob/1ab383bbc1764249a4320ca28ed3ac6682c8fbac/versions/G1R20_Fix4/src/vba/RPX_UI.bas#L511)、`reports/PRESET_ASSIGNMENTS.json`。これはコードの明示代入を抽出した値であり、Issue実行時の設定表全体を読み取った証拠ではありません。

したがって、今回の速度設定をPython成功手順やG20改善候補と同じ制御経路として評価することはできません。ただし、G20改善へ切り替えるだけで通常初期状態から2%を達成する、という検証結果もありません。

## 2. 成功形状の生成全体はVBAの通常経路に再現されていない

過去Pythonでは、既知形状の`xy/tri/material_id`だけを読み、材料・自重・境界条件を正規入力から作り直しました。当時の応力・速度・Fsを新しい求解のseedとして読んだ結果ではありません。しかし、形状は既存の履歴を利用しています。

- 下界：履歴付きnative1000 → uniform4000 → stress fan5350 → star5422 → 5594 → 5966。
- 上界：2914形状 → 5 passesの辺flip → 局所細分6106 → 方向性細分8739。
- 上界の初期2914形状までの手順と、一部の実際のmarkパラメータは完全には回復できていません。

Fix4のG20 balanced候補は、新しいTARGET1000形状から4分割します。上界flow候補は現在の形状の辺を組み替え、節点数・要素数を維持します。これらは成功した5966／8739形状を作る全履歴と同一ではありません。根拠：[RPX_Adapt.bas L2238以降](https://github.com/tak063495-prog/excel-vba-rpfem/blob/1ab383bbc1764249a4320ca28ed3ac6682c8fbac/versions/G1R20_Fix4/src/vba/RPX_Adapt.bas#L2238)、原記録`HISTORICAL_RECIPE.md`。

薄層・破壊機構を表現する近似空間の違いがGAPを残した可能性があります。要素数だけではその原因を確定できず、相互に包含しないメッシュの要素数増加は、必ずFsが改善するという保証にもなりません。通常の適応生成で成功形状へ到達する問題は未解決です。

## 3. 同じ原方程式でも、縮約・補助問題・fallbackが異なる

Python成功経路は、affine縮約後に解き、元空間へ復元して元q8問題で監査します。上界は原LP → cap付きQP → 同Fsでの再中心化 → 原dual helper → 自由等式乗数だけの線形補正を使用しました。QPのdualを原問題の根の証拠には使っていません。

VBAの直接DLL入口は、非縮約の原問題を要求します。また、`DIRECT_DLL=1`は各Fs点でDLLを先に試す設定で、監査された候補を採用できなければHSD等へ戻ります。帯LUの切替だけでも、HSDが一切走らなくなる設定でもありません。根拠：[RPX_Fs.bas L186以降](https://github.com/tak063495-prog/excel-vba-rpfem/blob/1ab383bbc1764249a4320ca28ed3ac6682c8fbac/versions/G1R20_Fix4/src/vba/RPX_Fs.bas#L186)、[RPX_AccuracyNative.bas L138以降](https://github.com/tak063495-prog/excel-vba-rpfem/blob/1ab383bbc1764249a4320ca28ed3ac6682c8fbac/versions/G1R20_Fix4/src/vba/RPX_AccuracyNative.bas#L138)。

この差は、同じ原物理式でも収束・所要時間・採否が同じにはならない理由です。物理式そのものが異なると確認した、という意味ではありません。Issue報告の19回のHSD時間切れ・約6522秒の内訳は報告者の集計に基づき、直接DLLの各棄却理由までは今回の資料から断定できません。

## 4. 未達時のown-q8再求解は速度面の確認対象

Fix4の終端証拠の再利用はGAP目標達成時に限られます。未達なら、選択した両側についてキャッシュと根をリセットし、FsEndpointを新たに実行します。根拠：[RPX_Adapt.bas L1457以降](https://github.com/tak063495-prog/excel-vba-rpfem/blob/1ab383bbc1764249a4320ca28ed3ac6682c8fbac/versions/G1R20_Fix4/src/vba/RPX_Adapt.bas#L1457)、同L1532–1539・L1588–1593。

報告された約4370秒のown-q8再計算と整合する経路はあります。実際に省略できるかは、同一メッシュ・モデル・q・方針・入力・両根の証拠が既に揃っていることの照合が必要です。GAP未達という理由だけで未監査の結果を再利用する提案ではありません。また、重複削減だけではメッシュ由来の30.5%を2%に縮めたことにはなりません。

## 次に切り分ける順序

1. 今回の設定表、各側のメッシュ・q・自己根・採用stage、各Fs試行のdirect採否とHSD時間を揃える。
2. 同じ正規入力でSpeed/G20を比較し、成功形状を使わない通常生成の到達範囲を分ける。
3. 既知形状では同じFsの原行列・物理場を比較し、その後に各自己根まで比較する。以前のVBA固定点一致は、この先頭部分までの検証です。
4. 支配要因を確認した変更だけを試し、原物理監査・根幅条件・GAP2%を保持する。

Issue #28は「HSD OPTIMAL後の別の復元失敗」、#29は報告上「HSD時間切れを含む適応経路の未達」です。同じ原因と決めつけず、[D05AのPython解析](../issue28_d05a_python_analysis_20261017/README_JA.md)も併せて検討します。
