# 水位適合VBA overlayの独立レビューと入力保護修正

対象は隔離試作 `versions/G1R19_WaterCommon_20261016`。実行中のG1R19_2pct、Excel、DLL、旧Python計算は変更していない。実Excelコンパイル・実行はintegration担当の確認待ち。

## 数式・幾何・境界

`RPX_Mesh`、`RPX_Water`をG1R19基準と改行正規化して比較した。既存関数の変更は `RPX_AssignRegions` の水位分割フックと `RPX_WaterApply` の適格な共通荷重経路のみ。旧材料・強度則・物理拘束・q・監査閾値は変更していない。

水平mode2、coupling ON、suction OFF、正の有限γw、等しい水位、完了したmode2結果の全節点水頭が同じ水位、土のみという条件を確認した。共有辺ごとの交点を1個にし、元の節点と材料を保持し、子要素を元要素内に分割して面積・一次モーメント・外部境界6フィールドを照合してから採用する。

水位線は `RConstraints` の幾何保護だけに追加され、`DFaces` や物理外部面、材料界面、速度接着、力学拘束を追加しない。水圧は `u=γw max(level-y,0)` の同じ式の厳密微分として、水位下で `-grad u=(0,γw)`、上で0を用いる。重力と水圧を各呼出しで作り直すので二重加算はない。重力の固定/基準、reference_total、地震荷重の既存分岐も保持している。

独立した100個のランダムメッシュ（元4906要素→子7704要素）で、元配列不変、共有交点、正の面積、材料保持、水位非横断、幾何制約と物理境界の分離を追加確認した。面積誤差最大5.685×10⁻¹⁴、一次モーメント誤差1.365×10⁻¹²。別の直接多角形積分との水位下面積誤差は2.843×10⁻¹⁴。これはVBA実行の検証ではなく、独立スカラー幾何検査である。

## 発見した入力保護の欠落と修正

旧 `RPX_Guard` の入力snapshotは要素定義・材料・設定のみで、水理入力/結果を含まなかった。試作の `WaterCommonLevel` はそれらのliveセルを再読込するため、solverのDoEvents中の水位・γw等の変更が入力epochに結び付かず、後の評価や確定結果が異なる水理sourceを使う危険があった。

親の指示により、隔離overlayに `RPX_Guard` を追加し、標準InputBegin/AssertInputs/InputEndへ水理sourceのbind/assert/resetだけを差し込んだ。適格な入力は既存の入力epochに固定する。前回のメッシュ/剛体状態に開始時の適格性判定を依存させず、実際のメッシュ適格性は従来通り別途検査する。

値とFormulaをVariant型・長さ付き文字列・Doubleの生IEEE754 hexで固定する。対象はlayout、mode、coupling、suction、γw、水位表全体、結果の収束/計算ID/mode/個数、節点座標/水頭、要素接続。元sourceの変更、欠落、epoch不一致は `WATER_COMMON_SOURCE_CHANGED` で中断し、mesh/resultを非currentにし、prepared状態・lower cache・Fs根を消す。開始時に不適格なら同一run中に共通水圧経路を後から有効化しない。一度共通荷重を適用したrunで適合性が失われても、旧P1荷重へ戻らず中断する。End/新Beginでsnapshotを解除し、previewへ状態を持ち越さない。

同じ一定水頭を持つ歴史的mode2結果は、同じグローバル圧力式を表す。適格性判定だけで過去の全入力変更に対する結果の新鮮さを証明したとは言えないため、READMEのstale表現を限定した。run開始後は上記の厳密source snapshotで変更を拒否する。

## 検証済み範囲と残る実VBA確認

幾何7テスト、入力保護契約/静的hook8テスト、復元失敗契約/静的handler6テスト、合計21件が合格した。入力保護テストは型違い、生Doubleの微差、±0、各source項目の変更、epoch、欠落、End/新run、preview、再帰防止、fallback外への例外伝播を検査する。Guardの既存本文は追加hook行を取り除くと基準ソースに完全一致する。

実Excelでは全3モジュールを一緒にimportしてコンパイルし、`tests/WaterCommonSourceProbe.bas` の11種類のmutation probeを確認する必要がある。加えて、初期水位分割とrefine後の荷重/境界、変更中断、幾何処理失敗時の復元、全原q8場監査・own rootを確認してから数値証拠として扱う。静的レビューだけで新VBAの2%達成や実際の例外動作が合格したとはしていない。

Meshの復元失敗を抑制していた箇所も、親の指示で隔離試作だけ修正した。別のerror handlerを持つ `WCRestoreMesh` へ元配列/辞書の復元、Topology/MapFaces、元境界の照合を分離した。どれかが失敗すると `WATER_COMMON_ROLLBACK_FAILED` に元エラーと復元エラーの番号・source・messageを別々に残し、必ず中断して旧経路へ続けない。元の取消または復元中の取消ならerror18を保持する。入力変更・取消は復元成功時も伝播する。失敗注入は既存TEST_INJECTION=1かつ明示armが必要な仕組みだけを用いる。

`tests/WaterCommonRollbackProbe.bas` に実VBA用8ケースを追加した。毎回、新しい非適合raw seedから開始し、復元に失敗したmeshを再利用しない。実際のhandler動作はExcelでの確認待ちであり、Pythonの静的契約テストをその代替証拠にはしていない。

再現:

```powershell
D:/Work/rpfem-venv/Scripts/python.exe -X utf8 -m unittest discover -s versions/G1R19_WaterCommon_20261016/tests -p "test_water_common*.py" -v
D:/Work/rpfem-venv/Scripts/python.exe -X utf8 validation/G1R19_WaterCommon_review_20261016/audit_water_common_static.py
```

正確なsource/import/baseのSHAはoverlayの `SOURCE_MANIFEST.json`、独立追加検査の結果は本フォルダの `STATIC_REVIEW_CHECKS.json` に保存した。
