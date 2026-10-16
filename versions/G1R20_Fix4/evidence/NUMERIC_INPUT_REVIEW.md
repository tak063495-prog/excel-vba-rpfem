# 数値・モデル入力レビュー

対象は公開G1R20 Fix3（main `d9d68d9971cc36c1eb012ac5bccf52813813f433`）。
旧版は書き換えず、所有する使い捨てExcelプロセスで公開ブックの挙動を調べた。

## 実証した問題

1. **地震荷重の固定／比例分類が前のFs解析に依存する。**
   `RPX_AssignRegions`でrawに戻す前に`RPX_SeismicBody`を呼んでいた。
   面積1、γ=20、kh=.2、自重固定の同じ入力に対して、通常`RPX_Run`の支持力が
   下界183.043478224／上界184.896505006となった後、実Fs解析を挟むと
   下界48.343139050／上界51.827715843となる。
   後者だけ水平体積力4が比例荷重へ移り、荷重倍率λに伴って4λへ増加した。
   入力ハッシュ一致、両方ともLastError空・ResultCurrent=Trueを確認した。
   最終負例は`results/seismic_negative_/1792183772543388000/summary.json`。
2. **共線辺を共有する部分重複領域を受理する。**
   異材料の矩形[0,1]×[0,1]と[.5,1.5]×[0,1]を受理する。
   .001幅の狭い重複を左右から作った場合も同様。
   既存のproper交差・頂点判定だけでは、辺上にある重複頂点を見逃していた。
   全定義点で既に分割済みの各辺について中点も分類する。
   隣接・内包・外周へ接する内包は受理し、完全同一・自己交差は拒否する規約を保持する。
3. **IDが空の部分行、表末尾を超える行を無言で除外する。**
   5表すべてで確認した。面IDだけを消して比例荷重-1を残すと、面数3→2、
   比例鉛直荷重合計-1→0として受理する。点1007行／材料1005行なども無言で無視する。
   UI／Guard側で対処する。
4. **物理数値のBooleanを暗黙に数値へ変換する。**
   TRUEの面荷重が-1となり、FALSEのc・γ・ψが0として受理される。
   負座標、c=0、φ=0、γ=0自体は有効なので、値ではなく型を区別して拒否する。
5. **旧テキストモデルの型・範囲・参照検査不足とファイル保持。**
   `Input#`でLongへ直接代入するため、0.4/1.4/2.4の節点indexが0/1/2へ丸まる。
   負c、未知の境界kind、内部辺への境界指定も受理する。
   EOFエラー62後、同じファイルの独占Openが55となり、読込ファイルが開いたまま残る。
   Variantで受けてから型・有限性・厳密整数・物性・参照を検査し、全失敗経路で自分の
   ファイルだけを閉じる。Live配列を置換するため、旧結果/current/certificateも失効させる。

入力行列と旧形式の最初の包括負例は
`results/seismic_negative_/1792183665675121600/summary.json`。
正常11、不正拒否28、既知の不正受理18、旧形式7件を含む。
元ソルバー・強度式・原q8監査・根の判定基準・DLLは変更していない。

## 独立した幾何参照

`tests/test_region_overlap.py`は、矩形1,350組を区間の交差・包含による解析的oracleと比較する。
凹L形と矩形900組は、独立した厳密な格子占有集合で比較する。
改修後判定は全2,250組で一致し、旧判定の取りこぼしもテストで確認する。
これはVBA実行の代わりではなく、別途行うExcel入力行列を補う数値参照である。

## 試験実装の修正と未確定のExcel終了

初期試行は省略せず保存した。UTFエンコーディング名の誤記、浮動小数点の完全一致比較、
引数なし`RPX_Setup:`を呼出しではなくラベルにしてしまった試験コードを修正している。
最後のfixtureは初回だけ明示`Call RPX_Setup`、以降はψ列とoverflow末尾を含め入力表を明示クリアし、
同じ入力を復元する。

実TestFsの次に毎回publicSetupを実行した別試行
`results/seismic_negative_/1792183690937194400`では、所有Excel15952が終了し、
回復子5400が`seismic (version 1).xlsb`を回復表示した。
`owned_recovery_child.json`に親子の作成時刻・コマンド・窓名と、確認後の終了を記録した。
一般COM例外だけでなく回復起動を確認したため、単なる試験失敗として扱わず、
別担当がpublicSetup／DLL／描画の切り分けを行う。
この時点で原因は断定していない。

## 実行方法

```
python -X utf8 versions/G1R20_Fix4/tests/run_seismic_phase.py --input-matrix --legacy
python -X utf8 versions/G1R20_Fix4/tests/run_seismic_phase.py --production
python -X utf8 versions/G1R20_Fix4/tests/run_seismic_phase.py --candidate --input-matrix --legacy
python -X utf8 versions/G1R20_Fix4/tests/run_seismic_phase.py --candidate --production
```

無指定は公開Fix3ブックそのもの、`--candidate`は使い捨てコピーへFix4全標準モジュールと
ThisWorkbookを入れてコンパイルする。試験は場・根・ソルバー状態を注入しない。
実行中の別Excelは拒否し、PID・作成時刻・コマンドを照合して自分のプロセスだけを扱う。
最終positiveの結果は検証完了後に別JSONへ集約する。

配布保存ブックの確認には `--final-book --expected-book-sha <SHA256>` を用いる。
この経路は保存ブックと同じ単一DLLを使い捨てディレクトリへ複製し、試験用プローブだけを追加する。
製品モジュールは差し替えない。main入力57件、legacy入力31件、実行中読込の拒否1件を別プロセスで、
q2/q8荷重位相と実public計算の履歴回帰をもう1プロセスで確認する。
両プロセスで保存ブック・製品ソースのハッシュ不変、所有Excelの終了と回復子の有無を記録する。

## 最終保存ブックの実行結果

`b010ab89c4e70540277475d1aab025db2545a24fd6697ad87856998737e94be6` の配布保存ブックで
`--final-book`を用い、製品モジュールを差し替えずに次の2プロセスを実行した。
集約は`evidence/NUMERIC_INPUT_EXECUTION_SUMMARY.json`、生記録は次の各`summary.json`にある。

- `results/seismic_candidate_/1792185516506912600/summary.json`: main 57件（正常11／不正46）、
  legacy 31件（正常3／不正28）、実行中読込の拒否1件、計89件が期待どおり。
  途中EOFを含む全legacy入力でファイルを解放し、待機時の読込では旧currentを失効させ、
  実行中読込ではcurrentと実行状態を維持した。
- `results/seismic_candidate_/1792185537367641200/summary.json`: q2／q8の固定・比例・kh=0、
  実Fsからの再メッシュ、実public bearing→Fs→bearingの計17呼出しがPASS。
  同一物理入力の固定荷重は初回／実Fs後／再作成で`(4,-20,0,0)`に一致した。
  支持力の下限`183.04347822419126`、上限`184.89650500619234`は実Fs前後で完全一致した。

両プロセスとも終了コード0、回復子0、ダイアログ0で、保存ブック、製品ソース、Fix3原本は不変。
独立幾何参照2,250配置を含むportable 7 testsも最終ソースでPASSした。
これは今回の不具合再現と回帰の確認であり、全26 benchmarkのGAP2%達成を示す試験ではない。

上記の単発Excel回復事象は追加のimmutable Fix3試験8 cycleでは再現せず、
切り分けと最終public再確認は`evidence/SETUP_CRASH_AUDIT.md`に記録する。
根本原因やDLL修正を断定していない。
