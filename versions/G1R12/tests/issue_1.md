# [P1] 適応解析の最終出力で bestPair メッシュと bestUpper 速度場が不一致になる

## 概要・優先度
P1: 適応解析の最終出力で、bestPair のメッシュと bestUpper の速度場が異なる候補に由来する場合、その組合せのまま上界場を書き出す経路があります。要素数が増えるケースでは配列範囲外、同じ要素数でも形状・接続が異なる場合には場とメッシュの誤対応につながります。

対象コミット: `890b1be5ef3e1ea72762ab7050647fdd52ecac01`（G1R11）

## 根拠となる呼出し経路
- [RPX_Adapt.bas L396–403](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Adapt.bas#L396-L403): 上界最小・下界最大の witness を独立に更新
- [RPX_Adapt.bas L308–311](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Adapt.bas#L308-L311): bestPair は pairGap 最小で選ぶため、bestUpper と同一候補とは限らない
- [RPX_Adapt.bas L1424–1441](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Adapt.bas#L1424-L1441): bestPair のメッシュを復元して最終候補を再評価した後、RestoreFinalBounds を呼ぶ
- [RPX_Adapt.bas L595–605](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Adapt.bas#L595-L605): 最良値と RBestUpper.field を RUpperField に戻すが、上界 witness の幾何はここで復元しない
- [RPX_Control.bas L183–194](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Control.bas#L183-L194): 適応解析直後に現在のメッシュを RPX_WriteMesh / RPX_Draw へ渡す
- [RPX_Output.bas L11–15](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Output.bas#L11-L15): 現在の re を使って RUpperField(12 * e + k) を参照する。RBestUpper の復元は [L58–60](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Output.bas#L58-L60) の出力後にある

## 成立条件と論理的反例
適応候補として以下の監査済み状態が保存され、最終再評価でも順位が変わらない場合を考えます（数値実測ではなく、選択条件を示す論理的反例です。剛体数 RR=0）。
- A: 601要素、bounds [0.60, 0.71]。pairGap ≈ 0.16794
- B: 1000要素、bounds [0.65, 0.72]。pairGap ≈ 0.10219

bestPair は B、bestUpper は A になります。B を復元した後、RestoreFinalBounds により長さ 12×601=7212 の A の上界場だけが戻ると、出力ループは B の1000要素を処理します。e=601, k=0 の添字7212は A の最大添字7211を超えます。

要素数が同じでも、節点座標・接続が異なる候補なら長さ検査だけでは防げず、別メッシュ上の速度を対応付けてしまいます。

## 影響
- 完了した適応解析の出力中に配列範囲外エラーとなり得る
- [RPX_Control.bas L283–290](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Control.bas#L283-L290) の失敗処理では bounds が消去されるため、保存済みの監査済み結果を最終結果として提示できなくなる可能性がある
- サイズが一致する場合は誤対応を検出せず、速度場・メッシュ表示を誤解させる可能性がある

## 修正案
1. 上界 witness のメッシュ、接続、材料・境界情報、場を同一 owner の一組として、RPX_WriteMesh / RPX_Draw / 上界場出力より前に復元する
2. 下界 witness は独立した owner のメッシュと応力場として出力する
3. 場長だけでなく witness のメッシュ識別・接続との一致を出力前に検証する
4. 復元位置を RPX_WriteFields の末尾から移すだけでなく、それより前のメッシュ書出し・描画も同じ owner を使うことを確認する

## 推奨回帰試験
- 上記 A/B の異なる要素数の状態を注入し、bestPair≠bestUpper でも範囲外参照せず、上界メッシュが A、上界場も A になる
- 同じ要素数で異なる座標・接続の2候補を使い、誤対応が起きない
- bestPair=bestUpper の通常経路を維持する
- bestLower が別候補の場合にも上下界それぞれの owner が維持される
- 最終候補の評価失敗後に保存 witness を使う経路でも整合が保たれる

## 確認範囲・未実施
上記は固定コミットのソースと呼出し経路を静的に確認した指摘です。Excelでの実行・反例の実測・修正後の回帰試験は未実施です。重点箇所のレビューであり、全機能や解析精度を認証するものではありません。

