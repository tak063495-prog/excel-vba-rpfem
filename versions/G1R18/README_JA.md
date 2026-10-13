# G1R18 精度改善実装

P06A・P06N の誤差1%目標は継続課題です。[Issue #11](https://github.com/tak063495-prog/excel-vba-rpfem/issues/11) に残しています。本版は精度改善の VBA 実装と追加 DLL の配布版であり、P06 達成版ではありません。Python の14/16ケース検証と、VBA/DLLの小規模回帰・場の移植検証は区別します。

## 利用

`delivery/RPFEM_20261014_G1R18.zip` を展開し、`RPFEM_20261014_G1R18.xlsm` を64ビット版 Excel で開いてください。2個の DLL はブックと同じフォルダーに置きます。Excel の通常のマクロ許可方法を使用します。Python は配布ブックの実行に不要です。

初期状態は従来経路で、追加回復は無効です。「設定」の「精度重視の設定」ボタンで Fs、q8、誤差目標1%、独立保存場の再監査、追加細分化を設定できます。この設定は長い計算を伴います。解析政策・材料・領域・荷重はユーザーの入力を保持します。

| 設定 | 通常 | 精度重視プリセット |
| --- | --- | --- |
| `ACCURACY_DLL_RECOVERY` | 0 | 1 |
| `QUADRATURE` / `ADAPT_CERT_QUADRATURE` | 既存 / 0 | 8 / 8 |
| `FS_TOL` / `GAP` | 0.0001 / 0.01 | 同じ |
| `TARGET` / `ADAPT_ACCURACY_CAP` | 既存 / 0（目標まで） | 1000 / 24000 |
| `CYCLES` / `ADAPT_ACCURACY_ROUNDS` | 既存 / 4 | 8 / 8 |
| `ADAPT_LOWER_TARGET` / `ADAPT_LOWER_FAN_ROUNDS` | 0 / 0 | 4000 / 1 |
| 既存解法 / 追加候補 / 反対側の時間 [秒] | 120 / 120 / 20 | 300 / 120 / 20 |
| `LOWER_FIXED_RECOVERY` | 0 | 0 |

追加回復の対象は元の自由荷重・q8・剛体なし・全参照荷重の上下問題です。対象外または DLL の利用に失敗した場合、既存解法と保存証拠を保持します。固定荷重の別回復は `LOWER_FIXED_RECOVERY=1` で個別に利用できます。

## 実装

* 粗段の Fs 許容幅を0.005から0.001へ変更し、最終0.0001を保持。上限と回数を指定した追加適応細分化、停滞・数値失敗時の復元を実装。
* 下界の一様細分化と c0 自由境界の星型分割。上界・下界は別メッシュと元の変数配置で保存。
* q8 精度プロファイルでは各保存側を元のメッシュで再計算し、元物理式と独立した根を再検証。他の q の根を q8 の根として扱わない。
* 条件を満たす六頂点の単一 c0 重力斜面には、任意の `ADAPT_C0_WEDGE_DEPTH>0` による薄い領域の上界候補。対象条件を満たさなければ保存場を保持。
* `rpfem_accuracy_x64.dll` は固定した Clarabel 0.11.1 を使用。元の疎行列をコピーし、上界の最小変更 QP（目的上限0.999998）と下界の自由λ QP（床1.000003）を独立に解く。
* 下界は元の同次等式に限り正倍率でλ=1.0000015へ縮小し、原式の平衡補正候補を LSMR 100/500/1000回まで試す。各候補を元の全等式・降伏・表面力・仕事で再監査する。
* 根の反対側には別の原双対実行可能性問題を使用。元の非縮約行列で stationarity≤1e-7、cone≤1e-8、厳密な目的分離を確認。QP・固定荷重・補助問題の双対値は根に使わない。
* 上界の保存候補は値≤0.999999、下界はλ≥1.000001。既存の原式監査1e-7を保持。根の完成には各元メッシュの幅≤0.0001×max(1,中点)とモデル・メッシュ・入力epoch・q・政策・端点・目標の一致を要求。
* 準備済みプログラムの Fs、正規化前後の目的係数、raw 等式、キャッシュ、変数配置を保存。古い Fs の compiled cone を別の Fs としてエクスポートしない。
* Issue #8 の水理残差修正を保持。

既存 `rpfem_native_x64.dll` の ABI `0x00170005` は保持し、新 DLL を並置しています。新 DLL の非同期 worker は入力を所有コピーし、キャンセル・結果取得・破棄を明示します。solver status は監査を代替しません。

候補の既定はα=0.01、平衡化OFF、thread1。独立双対は平衡化ON。`ACCURACY_DLL_CONE_SCALING=1` は各円錐の全行とオフセットを共通の正係数で正規化します。元行列のコピーは変えません。P06 の全実験をブックの自動処理で再現したという主張はしていません。

## 検証と再生成

リポジトリの `tools/build_release.py` が保持した G1R17、G1R18/RayGate テンプレートと水理修正から62モジュールを生成します。`src/vba` は UTF-8、`import` は CP932。`SOURCE_MANIFEST.json` は生成時の実ファイルバイトを記録します。G1R18のGit属性は改行変換を無効にし、Git blob・実ファイル・ZIP内の同じバイトを保全します。

```
python versions/G1R18/tools/build_release.py
python versions/G1R18/tests/test_accuracy_native_excel.py
python versions/G1R18/tests/test_accuracy_native_excel.py --fs
python versions/G1R18/tests/test_adapt_excel.py --failure-only
```

テストは専用の所有 Excel を順次作成して閉じます。元の非縮約 CSR と全ベクトルを保存し、Python 側でも再計算します。初期化済みの配布ブックは再オープンしてコンパイルと精度プリセットを確認します。配布 ZIP の各格納物と両 DLL は manifest で照合します。小規模試験の成功を P06 の精度達成に置き換えません。

旧 G1R18 の P06A 生産試験は有効な E898/q8 の下界0.813829144464255、上界0.895546879084081、誤差9.5611%を保持して終了しました。数値失敗後の有効場保持を確認した baseline で、新 DLL の P06 生産実行とは異なります。

Rust ソース、固定 lockfile、ビルド手順、ライセンス、17件の独立 FFI 検証は `native/accuracy` と `tests/test_accuracy_native_ffi.py` にあります。新 DLL のビルドは Windows x64・Rust 1.85・MSVC を使用します。配布には LSMR/SciPy を含む第三者ライセンスを同梱します。
