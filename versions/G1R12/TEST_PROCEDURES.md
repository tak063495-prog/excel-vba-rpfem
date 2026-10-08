# G1R12 の試験再現

Python、numpy、scipy、mpmath、pywin32等は `requirements-validation.txt` を参照してください。通常のExcel解析にPythonは不要です。

## Python検査

配布フォルダーで `python -X utf8 tests/python_pretest.py` を実行します。所有者選択の反例、実装の出力順序、変更モジュールの限定、エラー文法を検査します。これだけではExcelコンパイル・求解の合格になりません。

## 実Excelの回帰

1. Windowsの実Excelと、既に利用可能なVBEアクセスを使います。G1R11原本を用意し、必要なら環境変数 `RPFEM_BASELINE_XLSM` に絶対パスを設定します。原本のハッシュは `source_identity.json` と一致する必要があります。GitHub全体を取得した場合は隣の `versions/G1R11/` も探索します。
2. `python -X utf8 tests/native_regression.py` を実行します。独立したExcelインスタンスと使い捨てコピーで旧新比較を行います。製品の求解カーネルは変更せず、試験コピーにだけ失敗注入・状態作成probeを追加します。
3. `python -X utf8 tests/audit_results.py` で取得した場の原式・50桁監査を行います。JSONのexportsの保存場所を移動した場合は、それに対応する保存場のパスも調整してください。
4. 完成G1R12 XLSMがフォルダー直下にある状態で、`python -X utf8 tests/build_and_workflow.py workflow` を実行し、実際のRPX_Runを旧新の小規模ケースで比較します。続いて `python -X utf8 tests/bearing_regression.py` で小規模支持力の通常・適応経路を確認します。

`build_and_workflow.py build` はG1R11から完成ブックを新規構築する処理です。既存の完成ファイルを上書きしません。再構築するなら別の作業フォルダーにソース・原本参照・テストを置き、完成ファイルがない状態で使います。

試験コピーで入力表をクリア・小規模ケースに差し替えます。原本・製品コピーは変更しません。処理が中断した場合は自身が起動した試験Excelだけを終了し、他のExcelやユーザーのブックを終了させないでください。

## 目視受入れと未実施の全ケース

- Issue #1: 上界Aの601要素と下界B/最良区間の1000要素を別々に保存し、最終上界の節点A:C・要素A:F・速度H:S・図がA、下界F:H・AA:AF・応力AH:AYがBであることを確認します。同数・異座標でも繰り返します。場やqを意図的に不一致にした場合は停止する必要があります。
- Issue #2: 実因子分解分岐の詳細エラーがログに残り、同一問題確認後に`HSD_RESCUE`、救済1回、独立監査合格になることを確認します。番号7/9/18、未知、監査失敗、モデル変更は救済しないことを確認します。
- 全P06N: 配布設定のまま解析を実行し、ログ末尾、根幅、物理監査、GAP、採用メッシュと場、実行時間を確認します。今回NOT_RUNです。
- 25ケース: `excel_acceptance_g1r12.csv`の各ケースを個別に実行し、入力・ログ・場・上下界の証拠がそろった行だけ更新します。間隙水圧のD08Aは対象外です。元ケース資料が必要で、入力台帳や以前のPython結果を今回のExcel実行PASSへ転記しません。
