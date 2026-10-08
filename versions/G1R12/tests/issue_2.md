# [P2] pivot情報付き分解失敗が RobustEligible に一致せず HSD 救済をスキップする

## 概要・優先度
P2: legacy ソルバーが実際に送出する `NEWTON_FACTORIZATION_FAILED pivot=...` が、RPX_RobustEligible の文字列完全一致に一致せず、意図された HSD 救済を試さずに再送出されます。

対象コミット: `890b1be5ef3e1ea72762ab7050647fdd52ecac01`（G1R11）

## 根拠となる呼出し経路
- [RPX_Solve.bas L409–421](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Solve.bas#L409-L421): FactorSystem の分解失敗時はエラー番号5で `NEWTON_FACTORIZATION_FAILED pivot=` にピボット情報を付加して送出
- [RPX_Solve.bas L626](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Solve.bas#L626): OptimizeCore から FactorSystem を呼び出す
- [RPX_Solve.bas L1762–1767](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Solve.bas#L1762-L1767): RPX_Optimize の失敗処理は元の errorText をそのまま再送出
- [RPX_Robust.bas L10–15](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Robust.bas#L10-L15): Select Case Trim$(message) の許可リストには接尾情報のない `NEWTON_FACTORIZATION_FAILED` のみ
- [RPX_Robust.bas L44–49](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Robust.bas#L44-L49): LegacyFailed → LegacyRejected でも正規化せず、適格性が False なら再送出
- 救済に入った場合にだけ [L50–64](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Robust.bas#L50-L64) の同一問題チェックと HSD 再試行が行われる

## 成立条件・最小の論理的反例
XCancel=False、RPX_AuditRejected=False、エラー番号5の場合:
- `RPX_RobustEligible(5, "NEWTON_FACTORIZATION_FAILED")` は True
- `RPX_RobustEligible(5, "NEWTON_FACTORIZATION_FAILED pivot=123")` は False

後者が FactorSystem の実際のメッセージ形式です。Trim$ は前後空白しか除去しないため、この相違は残ります。これはコードの分岐条件から導いた反例であり、Excelで実測した結果ではありません。

たとえば c>0 など、useHSD=False で legacy 経路を使う問題で分解が失敗すると、この理由だけで HSD_RESCUE に到達しません。初めから useHSD=True で HSD を呼ぶ経路とは区別が必要です。

## 影響
分解失敗を救済対象として列挙しているにもかかわらず、診断用の pivot 情報が付くことで救済機会を失い、その解析点が失敗として返ります。HSD を試せば必ず解けるという指摘ではなく、許可された再試行が実行されないという制御フローの問題です。

## 修正案
1. エラー種別を詳細メッセージから分離し、厳密なコード／トークンで許可リスト照合する
2. 当面文字列を使うなら、既知の区切りに基づくトークン抽出等で診断 suffix を分離する。[RPX_Guard.bas L333](https://github.com/tak063495-prog/excel-vba-rpfem/blob/890b1be5ef3e1ea72762ab7050647fdd52ecac01/versions/G1R11/src/RPX_Guard.bas#L333) の既存分類も参考にする
3. 元の詳細メッセージはログ・再送出用に保存し、pivot 情報を失わない
4. 単純な Contains 判定への変更は避ける。未知のエラー、キャンセル、監査拒否、別種の名前が部分一致して救済されないようにする
5. 同一モデル・入力世代・荷重段階・強度低減率・積分条件のチェックと、救済回数制限を維持する

## 推奨回帰試験
- bare token と `pivot=123` 付き token がともに適格となる（番号5、キャンセルなし、監査拒否なし）
- 未知のエラー、似た名前、前後に任意文を加えた文字列を誤って適格にしない
- 番号5以外、XCancel=True、RPX_AuditRejected=True はともに不適格のまま
- FactorSystem の失敗を注入し、実際の装飾メッセージが RPX_Optimize を経由したとき、同一問題に HSD_RESCUE が1回だけ入ることを確認
- HSD 側も失敗した場合はその失敗を正しく返し、成功扱いや無制限再試行にしない
- 通常成功・初回から HSD の経路に挙動変化がない

## 確認範囲・未実施
固定コミットのソースと呼出し経路の静的確認です。Excelでの実行、実際の分解失敗注入、修正後の回帰試験は未実施です。重点箇所のレビューであり、全機能や解析精度を認証するものではありません。

