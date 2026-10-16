# 初期化入口の実Excel検査

製品の変更は4箇所で引数なしSub呼出しに `Call` を明示したものです。既存結果の偽合格が確定したという主張ではありません。製品ソースは `versions/G1R19_StateReset_20261016`、テスト専用63モジュールは各結果フォルダ内に独立保存しています。

最終比較は次の2回で合格しました。どちらも専用Excelのみを使い、正常退出を確認しました。元ソルバー・DLL求解・水理FEM求解の前にテスト専用trapで停止しており、場・Fs根・精度の証拠を生成していません。

| 役割 | 結果フォルダ | 観測 |
|---|---|---|
| 修正版 | results/positive_/1792153293517482300 | Conic消去、Direct下・上private状態消去、Krylov取消入口、水理Defaultsの5検査にPASS |
| 修正前の呼び方 | results/negative_/1792153314267480200 | 同じ5検査が意図したsentinel/trap理由でFAILしたため、比較試験としてPASS |

最終StateReset manifestは `a26beb0538dfccd33021c4ea603a702957a2452b44ed08dc23bfbe75b088a147`、単一DLLは公開G1R19と同一の `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b` です。WaterCommon3モジュールも合成していますが、この試験の入力D02は乾燥条件です。

以前の失敗も残しています。最初のハーネスは設定行追加がなくSEISMIC_KHで停止し、次の比較用Krylovはテストtrapの位置が遅く配列初期化前にerror9で停止しました。いずれも求解前です。後者はテストtrapをHEnsureDD直前に移し、製品4モジュールのバイトを変えずに最終比較を実行しました。旧失敗を最終PASSに付け替えていません。遅延した所有Excel退出確認は `LATE_OWNED_EXIT.json` に分離保存しています。
