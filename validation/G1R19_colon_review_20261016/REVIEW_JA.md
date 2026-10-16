# VBA の先頭 SubName: 構文と証拠の寿命の独立確認

2026-10-16。`versions/G1R19` に `G1R19_2pct` の同名モジュールを重ねた62モジュールを読み取りだけで確認した。Excel起動、求解、製品ソース変更は行っていない。別担当の実Excel検証で判明した「物理行先頭の引数なし `SubName:` が呼び出しではなくラベルになる」という構文を対象とする。一般のコロン区切り文を全て無効とみなすものではない。

## 結論と最小修正候補

該当箇所は4件。先頭の名前に `Call ` を付ければ意図した呼び出しを明確にできる。実行中の凍結ソースには適用していない。

| 現在の位置 | 省略される処理 | 影響の判定 |
|---|---|---|
| `versions/G1R19/src/vba/RPX_Conic.bas:914` | `RPX_OriginalGuideClear` | Farkas候補ごとのClearが実行されない。現在の受理経路では下記の独立した寿命・同一性チェックが残り、これ単独による古い無効なdualの受理は確認できない。明示Callによる修正を推奨する。 |
| `versions/G1R19_2pct/src/vba/RPX_Fs.bas:189` | `RPX_ResetPreparedState` | 新規組立後のDirect求解前に、作業配列・warm state・originalKを失効させる意図が実行されない。続くCompileは実行され、原問題とbulk stateは更新されるが、Compile自体はSolveの全private cacheを失効しない。Direct失敗後のprepared=True継続を含むため、単なる表示問題ではなく状態再利用の修正候補。誤った最終物理証拠の受理は実証していない。baselineの同じ位置も該当する。 |
| `versions/G1R19/src/vba/RPX_Solve.bas:2073` | `HContext` | native Krylov入口で入力・epoch・現在iterate・factorの検査が一度省略される。各Stepの直前には独立したstandalone HContextがあり、検査が全面的に失われるわけではない。ただしNativeKrylovBeginより前の検査として明示Callが妥当。 |
| `versions/G1R19/src/vba/RPX_Seep.bas:433` | `RPX_SeepDefaults` | `RPX_SeepTestRelax` のテスト補助関数。通常の水理求解ルートではない。テストを前回の内部設定へ依存させるため修正候補。 |

4件とも同じモジュール内に同名へのGoTo/GoSub/Resume参照がなく、意図した分岐先ラベルと判断する根拠は見つからなかった。入力pinや証拠/cacheクリアを目的とする引数なしSubについて、これら以外の同じ先頭構文は62モジュール走査で見つからなかった。引数が全てOptionalのSubも検査対象に含めた。

Direct入口のReset省略については、現在のOptimizeCoreはwarm startを無効にしてColdPrimalDualを使い、KKT組立はoriginalKを上書きし、HSDは構造世代を進めて係数評価cacheを再構築する。したがって、残留フラグがそのまま誤った行列や場として採用されるとまでは結論しない。意図された失効処理が省略されること自体と、将来そのフラグを使う経路への影響を修正対象としている。

## Original Farkas guide が現在のコードで古い候補を受理しにくい根拠

1. `RPX_Solve.bas:1939` の各HSD入口は `RPX_P4Begin` でsolve epochを進める。`OptimizeHSDCore` はXStatusをRUNNINGにし、`RPX_HSDCandidate.bas:64` の実際に呼び出されるstandalone `RPX_OriginalGuideClear` を毎回通る。
2. `RPX_HSDObserve` は候補ごとにcandidate IDを進める。`TryHsdCandidate` はObserve→Build→Validate→Commitの順で、`RPX_OriginalFarkasGate` がFalseの候補はCommitへ進まない。Commit自体もvalid/numericalGate/physicalGateを要求する。
3. Gateは原full problemのzero cost、upper/q8、非affine、Fs、input/solve/geometry epoch、原問題の寸法、ExactModel、load phaseを検査し、原行normから戻したyと原coneのzを全変数で監査する。正規化後のstationarity<=1e-7、cone<=1e-8、separationの条件は変更されていない。
4. `RPX_Conic.bas:958` のCurrentはXStatusがORIGINAL_FARKAS_GUIDEであることに加え、ready、Fs、q8、input epoch、solve epoch、candidate ID、ExactModel、load phaseを再照合する。Fs側もCurrentがFalseならこのstatusを根の正符号に採用しない。
5. 原問題のCompile/Beginにも別のClear経路が残る。したがって、候補失敗後にreadyだけがTrueのまま残っても、通常ルートのID/status/epochの不一致で古い候補を排除する。

これは全ての外部macro呼び出しや将来の変更に対する完全性の証明ではない。公開Gateを同じcontext/IDのまま直接再呼び出しする場合には、失敗時にreadyが残るという局所的な契約違反がある。前の証拠が同一原問題で有効なら数学的有効性は失われないが、「今回の候補は拒否されたのでreadyもFalse」という意図を満たすため、先頭Clearを実呼び出しへ直すのが適切である。ガイドは離散根の符号のみで、速度場・外部荷重仕事・連続体上界として採用する経路は確認できない。

## code_key の範囲

`versions/G1R19_2pct/tools/build_candidate_book.py` の関数をASTから単独抽出し、Excelを使わず8比較を実行した。識別子の大文字小文字とCRLF/LF差は無視する一方、文字列（連続した二重引用符を含む）、アポストロフィコメント、数値・演算子・空白・Callの追加削除は区別する。Remコメントの大文字小文字は通常コードと同様に小文字化するので、「全コメントのbyte同一性を保護」と表現することはできない。

現在のbook builderはAddFromString直後のloaded codeと再open後のcodeを比較している。原ファイルのsource SHAは別に記録するが、`code_key(code(source))` とloaded keyの直接一致は検査していない。import時の意図しない変換も検出するなら、この比較を追加するのが小さな改善になる。今回の曖昧なSubName:は原文自体がラベル構文であり、保存時の大文字小文字変換だけで発生した問題とは断定しない。

## 再現

```powershell
D:/Work/rpfem-venv/Scripts/python.exe -X utf8 validation/G1R19_colon_review_20261016/audit_statement_colons.py --output validation/G1R19_colon_review_20261016/STATIC_AUDIT.json
```

JSONには62モジュール全byte SHA、該当行、引数なし/全OptionalのSub名の件数、同名への分岐参照、code_keyの8比較を保存する。これはソース静的監査であり、4行を直した候補ブックの実Excel compile/reopenや数値回帰に代わるものではない。
