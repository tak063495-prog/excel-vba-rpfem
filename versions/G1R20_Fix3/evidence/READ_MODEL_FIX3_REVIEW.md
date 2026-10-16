# 旧テキストモデル読込の制約集合初期化

## 確認した不具合

Fix2 の `RPX_ReadModel` は `RBoundary` を再生成する一方、
`RConstraints` を初期化しない。旧テキスト形式は制約線分のレコードを
含まないため、新規読込時の正しい集合は空である。

- 新規セッションでは `RConstraints Is Nothing` のまま、
  `RPX_ExactModel` の `RConstraints.count` で実行時エラー91となる。
  `RPX_TestFs` からも同じエラーになる。
- 別モデルの制約集合が存在すると、その集合が新しいファイルのモデルIDに混入する。
  読込直後に `9998:9999` という旧制約が残る負例で確認した。

## 最小修正

`RPX_Model.bas` の `Set RBoundary = RPX_Row()` の直後に
`Set RConstraints = RPX_Row()` を追加した。追加コメントを含む2行のみ。
数値式、物理監査、原q8、GAP、根幅、DLLは変更しない。

## 同時に確認した解析的Fsと直接回復の状態衝突

配布済みの改善プリセットは `DIRECT_DLL=1` である。全材料の摩擦角0のとき、
`RPX_FsEndpoint` は `FsValue(..., allowRecovery=False)` で厳密な最適化を要求するが、
Fix2ではこの指示が `RPX_Bound` の直接DLL経路に伝わっていなかった。
直接経路が片側の双対根拠を採用すると `ORIGINAL_ACCURACY_DUAL_GUIDE` になり、
続く解析的Fs枝の `OPTIMAL` 要求により `NO_FINITE_POSITIVE_FS` で誤拒否した。

`results/read_model_fixed_1792181286402870100/summary.json` では、上記c=10のモデルで
`RValue=XObjective=10.000000000000032`、`AccuracyDirectLastKind=original_dual`、
採用True、状態 `ORIGINAL_ACCURACY_DUAL_GUIDE` とエラーを同時に確認した。
この段階ではモデル読込初期化を直した6モジュール候補で、FsはFix2と同じである。

修正は `RPX_Bound` に既定Trueの `allowDirect` を追加し、直接経路の条件へ適用し、
`FsValue` から既存の `allowRecovery` を渡す3箇所のみ。
摩擦角0の解析的Fsで、片側根拠を最適解として扱わず、従来の最適化と物理監査を通す。
通常の直接DLL設定、物理基準、根幅、GAPは維持する。

## 回帰の対象

`tests/run_read_model_fix3.py` は配布Fix2ブックのコピーを新規の専用Excelへ開き、
負例にはFix2全63モジュール、正例にはFix3全63モジュールを読み込む。
旧配布ブック、旧ソース、既存証拠は変更しない。所有PIDの終了と旧ファイルの
SHA256不変を記録する。

正例では、初回読込、同一ファイルの再読込、別材料ファイルの読込、元ファイルへ戻る
連続読込、旧制約が存在する状態からの読込を確認する。
さらに2三角形の単位正方形で、下辺固定・他3辺に一定せん断の基準表面荷重を与える。
一定応力と速度 `vx=y, vy=0` が対応するため、摩擦角0・粘着力10/11のFsは
それぞれ約10/11となる。原q8とraw監査を有効にして、初回・別モデル・元モデルの
`RPX_TestFs` を連続実行する。

## 証拠

最終ハーネスでのFix2負例は
`results/read_model_before_1792181535091900600/summary.json`、
Fix3正例は `results/read_model_fixed_1792181547034306500/summary.json`。
両実行とも新規専用Excelの終了と旧ブック／旧ModelソースのSHA256不変を確認した。

正例では全63標準モジュールとThisWorkbookを最終Fix3ソースから読み込んだ。
初回、同一、別モデル、元モデル、旧制約ありの5読込は、すべて空の制約集合となり、
同じモデルのIDは一致し、粘着力を変更するとIDも変わることを確認した。
`DIRECT_DLL=1`、`ACCURACY_DLL_RECOVERY=1` を維持したまま、Fsは次の通りとなった。

| 連続呼出し | 下界Fs | 上界Fs |
|---|---:|---:|
| c=10 初回 | 9.999989784787447 | 10.000009993685573 |
| c=11 別モデル | 10.999988821899189 | 11.000010992573849 |
| c=10 に戻る | 9.999989784787447 | 10.000009993685573 |

3呼出しとも状態 `OPTIMAL`、q8、物理gate採用、正規化／raw監査≤1e-7、
根探索未完了False、直接回復の採用Falseを確認した。
DLLは各試験コピーと同じフォルダの `rpfem_native_x64.dll` 1本だけがロードされた。
これはDLL設定を無効にするテストではなく、解析的Fsの厳密な最適化要求が
直接回復の片側根拠へ置き換わらないことを確認するもの。

`tests/test_fs_exact_contract.py` のportable 3件もPASS。
引数伝播と通常経路の既定True、Fs全体の変更が上記3箇所のみであること、
せん断fixtureに独立した一致上下界の解が存在することを確認する。
結果・元証拠SHA256・最終Fs/ModelソースSHA256・DLLパスの短い集約は
`evidence/READ_MODEL_FIX3_SUMMARY.json` に保存した。

最終portable全体は `evidence/PORTABLE_FINAL.json` の125件PASS
（Fix3新規10件、継承115件）。先行の `PORTABLE_DUPLICATE_SUITE_ATTEMPT.json` と
`PORTABLE_PRELIMINARY.json` は、ハーネスのFix2行がFix3を重複実行し、Fix2を
実行していなかったため最終結果として扱わない。その経緯も削除せず保存した。
該当行をFix2へ訂正し、最終125件を再実行した。製品ソースには影響しない。
準備中の2実行では、テキストfixtureのLF終端と、失敗がDLL到達前なので未ロードとなる
ことに対するハーネスの期待値を調整した。これらの結果も削除せず保存している。
製品側の不具合の判定には、エラー91と旧制約残留を再現した負例を使用する。
