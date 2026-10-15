# G1R19：2% 基準・高速化・単一 DLL

G1R19 は、合格基準を対称上下界幅 **2% 以下**とし、検証済みの精度を維持する
高速化と DLL の統合を行う版です。P06 の 1% を狙う追加精度実験は、
2026-10-15 のユーザー判断により終了しました。

## 導入

[G1R19 Release](https://github.com/tak063495-prog/excel-vba-rpfem/releases/tag/G1R19)
の `RPFEM_20261015_G1R19.zip` を展開し、64bit Excel で
`RPFEM_20261015_G1R19.xlsm` を開きます。
**同じフォルダーに置く DLL は `rpfem_native_x64.dll` の1個だけです。**
G1R19 のブックと DLL を組み合わせて使用してください。通常の実行に Python は不要です。

「設定」の「2%・速度優先の設定」で、目標幅2%、原q8監査、DLL先行を設定できます。
材料・領域・荷重・解析方針の入力を保持します。配布ブックは未解析の状態です。

| 設定 | 2%・速度優先 |
|---|---:|
| `GAP` | 0.02 |
| `QUADRATURE` / `ADAPT_CERT_QUADRATURE` | 8 / 8 |
| `FS_TOL` | 0.0001 |
| `DIRECT_DLL` / `ACCURACY_DLL_RECOVERY` | 1 / 1 |
| `ADAPT_PRECISION_FROM_START` | 0 |
| `RAW_LOWER_AUDIT` / `C0_AUTO` | 1 / 0 |

粗い探索段階の設定と、最終q8監査は区別します。原物理残差・根幅・目標幅の判定は
別々に維持し、物理監査の誤差基準を緩めて高速化する変更はしていません。

## 変更

- 従来演算の C/C++/MASM と、精度回復の Rust/Clarabel を静的に統合。
  従来94関数と精度回復9関数を同じ DLL から公開します。
  別の精度回復 DLL への転送や、実行時の2個目の DLL は不要です。
- 対応する q8 原問題は DLL で先に解き、原物理場または原双対の厳密検査を通った場合に
  従来の反復計算を省略します。対象外・検査不合格・DLL利用不可の場合は従来経路へ戻ります。
  ソルバーの status だけで場や根の探索符号を採用しません。
- 同じ計算内の q8 保存場・モデル・メッシュ・入力世代・方針・根区間が一致し、
  目標幅を満たすペアは、両側を元の物理式で再監査して最終再探索を省略します。
  条件が異なる場合は従来の再探索を実行します。
- G1R18 公開後の Issue 12（下界キャッシュの積分次数）と Issue 15
  （浸透結果の無効化）を取り込んでいます。

## P06 の精度

| ケース | 下界 Fs | 上界 Fs | 対称幅 | 2% 基準 |
|---|---:|---:|---:|---|
| P06A | 0.8468 | 0.85755 | 1.2614779828% | 合格 |
| P06N | 0.678279841819045 | 0.6876 | 1.3647112865% | 合格 |

幅は `2*(U-L)/(U+L)`。
これは[保存済み Python 場の独立 q8 再監査](https://github.com/tak063495-prog/excel-vba-rpfem/blob/main/validation/accuracy_20261013/snapshots/20261014_progress7_audit/REPORT_JA.md)
による確定値です。G1R19 が初期状態からの本番解析で同じ精度を自動再現したという
意味ではありません。P06N は既存 Davis 等価関連モデルの `hold_initial` 方針です。
今回の停止した追加精度実験から新しい上下界は採用していません。

## 検証・ソース

検証結果と速度の計測条件は [TEST_RESULTS.md](TEST_RESULTS.md)、ハッシュは `SHA256SUMS.txt` に記録しています。
DLL の両 ABI、従来の演算・メッシュ・取消・精度回復、実 Excel の採用と復帰、
入力が異なる場合の再利用拒否、DLL 1個のみのロードを確認しました。
小規模回帰と本番規模の初期状態からの精度再現は区別します。

- [統合DLLのソース・ビルド・ライセンス・headless回帰](native/unified/README.md)
- [VBAソース](src/vba) / [CP932インポート用](import) / `SOURCE_MANIFEST.json`
- [方針変更・実験停止の記録](https://github.com/tak063495-prog/excel-vba-rpfem/blob/main/validation/accuracy_20261015/README.md)
- [P06の本番確認などの課題](https://github.com/tak063495-prog/excel-vba-rpfem/issues/11)

ビルド・回帰の再実行にはリポジトリ全体の checkout を使用します。
公開済み G1R18 とその配布 ZIP・数値証拠は過去版として保持しています。
