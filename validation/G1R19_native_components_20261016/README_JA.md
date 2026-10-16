# G1R20へ取り込む既完了のVBA検証

水理・初期化・メッシュ候補・D02の検証記録を、ソース構成ごとに保持した証拠です。
P06のVBA完走や全26例のVBA完走を表すものではありません。

| 構成・対象 | 確認結果 |
|---|---|
| 7修正モジュール・D08A新規計算 | 883.952秒、GAP1.86070015%、両q8・両自己根・独立物理監査合格 |
| 水理入力固定・復元micro | 入力変更11種、復元8種、幾何・荷重8段階。求解0 |
| 明示Callの初期化micro | 修正5入口PASS、未修正5入口は意図した失敗。求解0 |
| 10修正モジュール・D02A新規計算 | 148.337秒、GAP0.00343004%、両q8・両自己根・36原CSR再生合格 |
| FlowMesh実Excel | A/N/N追加保護辺の3形状でPythonと完全一致、773/362/354辺変更。求解0 |

各専用Excelの終了を確認しています。D08とD02の時間は、異なる同時実行負荷での1回の測定です。
単一DLLのSHA256はG1R19と同じ `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b` です。

`evidence/COMPLETED_COMPONENT_PROOFS.json` に733ファイル・54,058,661 byteの証拠ZIPのSHAを記載しています。
原ブック、元ソース、入力、場、literal原問題、監査、実行ログは元のbyteを変えずにZIPへ保管しました。
ZIPはG1R20のRelease asset `COMPLETED_COMPONENT_PROOFS.zip` として公開します。
`package_native_evidence.py --output <新規ZIP>` で同じ対象を再収集できます。
