# CPU差に対応する再監査adapterの独立レビュー

確認日：2026-10-17（日本時間）。対象は今回追加した `portable_program_audit.py`、現在の `VERIFY_MICRO.py`／`VERIFY_HISTORICAL.py`、参照行列と各manifest、`TEST_PORTABLE_AUDIT.py`、README／ANALYSISのCPU差に関する説明である。

**高確信の不具合や原物理・原dualの合格条件を弱める変更は見つからなかった。** この判定は読取りによるソース・数式・SHA・構文のレビューであり、このreviewerが追加の求解、Excel、第三のCPU別再監査を実行したという意味ではない。

## 係数の比較と原門は分離されている

- 保存された参照行列は、それぞれ過去のproducerが記録した原行列のraw SHAと完全一致することを要求する。その参照NPZ自身のサイズ・SHAもmanifestで検査する。単に現在の行列に近いものを参照として採用する実装ではない。
- 再組立行列との比較では、CSCのshape・pointer・indexとその順序、dtype、整数cone構成を完全一致させる。有限のfloat64係数に対してのみ、配列単位の `32*eps*max(1,max(abs(reference_array)))` の明示幅を使う。実際の差、幅、raw SHAの一致／不一致を結果に残す。
- これは係数同一性の比較である。coreの原物理監査、raw stationarity `1e-7`、dual cone `1e-8`、原dual目的値・witness条件を相対化する変更ではない。
- MICROはfresh再組立と保存参照の両問題を元の `core.evaluate` で検査し、physical／witness／numericalの各合否を元記録と完全一致させる。fresh問題の原dual guideも元の門で再検査する。scalar再監査値の比較は明示された狭い丸め幅で行うが、合否のbooleanを許容差で変える処理ではない。
- HISTORICALもfresh再組立と保存参照の両方で元の数値門の合否を要求する。attempt65の原dual不合格を維持し、物理場の合格と混同しない。原full dualの逆変換のバイト一致照合も残っている。

## 過去のバイトと現在のadapterを区別する結合

旧 `HISTORICAL_MANIFEST.json` は変更せず、そこに記録された旧verifierの原本を `historical/program_refs/ORIGINAL_VERIFY_HISTORICAL.py` から読取り、旧SHAで照合する。現在のadapter、共通比較helper、追加参照行列は新しいreference manifestで別に結合する。この読取用の移動がコードとREADMEに明示されている。

peerの読取り照合で、旧72必須ファイル（旧verifier原本の移動先を含む）のサイズ・SHA不一致は0件だった。新しい歴史的adapter/reference manifestの5ファイルも不一致0件。MICRO側の2参照NPZと、凍結された6producer結果のサイズ／SHA・原行列digestの結合も一致した。元producer・元数値ソース・入力・過去ベクトルを書換えてCPU差へ合わせた記録にはなっていない。

MICROは6つの必要routeが欠落・重複なく存在すること、entry名とcase／routeの対応、Fs／q、再構築したmesh/model/policyを検査する。旧参照のdigestにも現在の厳密な再組立digestにも一致しない第三のdigestは、その係数証拠がないものとして拒否する。

## 負例と検証範囲

`TEST_PORTABLE_AUDIT.py` は、明示係数幅を超える改変、CSC index、dtype、NaN、cone構成、未知digest、欠落／重複route、q不一致、case／route不整合、参照NPZの1byte改変を拒否する内容だった。1 ULPの係数差を許す正例と、元ファイルを実際には書換えないNPZ改変の負例もある。peerはこのテストを再実行せず、コードと構文を確認した。

README／ANALYSISは、初回CIの行列SHA差を今回のPython再監査の移植性の問題と説明し、Issue #28のVBA原dual回復失敗の証明とは扱っていない。また、係数比較の丸め幅と原物理／原dualの受入門を区別している。この説明はレビューしたadapterと整合する。

この比較幅は、あらゆるCPU・BLAS実装で成功する数学的保証ではない。新環境で差が幅を超える場合、構造が変わる場合、元の物理・dual合否が変わる場合には拒否する。特定CPU経路での実行結果と最終ZIP展開での検証結果は、別の実行記録を参照する。P06Nの通常cold 2%達成、D05AのFix4修正完了、新規FOS根・2%完走を示す変更ではない。

## 読取り時のソースSHA256

| ファイル | SHA256 |
|---|---|
| `portable_program_audit.py` | `5204f04b2b6aba60f39d839e0f03576c9bfa5dde60b283adc9525eefe633ac90` |
| `VERIFY_HISTORICAL.py` | `aae0248c84dff415500e904e1f3fc69125ffaf2bd8ccab69bd9f5d13b511a757` |
| `VERIFY_MICRO.py` | `7c95a4815166ec973a933cf1363b2aaf1337741bd124ba23f35bf92bf6fa2c88` |
| `TEST_PORTABLE_AUDIT.py` | `23ad7164438e8e3b9457171c247e627fc9d7939949c0d42c682c4e23f913b557` |

4スクリプトの構文検査は合格。製品VBA・DLL、既存reports、元ソース、元manifest、保存ベクトル、Gitはpeerでは変更していない。この新しいreview文書だけを追加した。
