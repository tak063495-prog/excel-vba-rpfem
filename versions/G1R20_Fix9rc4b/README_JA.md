# G1R20 Fix9 rc4b（候補）

Fix9 をもとにした候補版です。変更は `RPX_Adapt.bas` だけです。コードは Codex が書き、こちらで確認しました。

## 変わったこと（Fix9 との違い）

- **薄層チェーンの早期開始**（`ADAPT_THIN_EARLY_START`、既定 1）：薄い材料層を検出したとき、条件を満たせば通常の細分を飛ばし、粗いメッシュの q8 の組からすぐ薄層チェーンを始めます。条件は、上界 10 段以上、下界 3 段以上、上限／粗い要素数が 2 倍以上などです。2 にすると常に早期開始、0 にすると Fix9 と同じ流れです。判定は `thin_layer_early_start` に記録します。
- **早期開始したときだけ**、上界チェーンを θ 0.95・要素数上限 7000 にします（`ADAPT_THIN_EARLY_UPPER_THETA`、`ADAPT_THIN_EARLY_UPPER_CAP`）。早期開始しないときは、これまでの値（θ 0.8・上限 3500）のままです。
- 下界チェーンの要素数上限は 6000（`ADAPT_THIN_LOWER_CAP`）です。
- 補助の上界解は、状態 AlmostSolved も指標に使います（`ADAPT_THIN_AUX_ALMOST=1`）。
- 精度の規則（q8、FS_TOL、監査、根の決め方、証明書）は変えていません。

## 配布物（`delivery/`）

| ファイル | 内容 |
|---|---|
| `RPFEM_20261019_G1R20_Fix9rc4b.xlsm` | ブック（VBA を入れ替え、コンパイル・開き直しの確認 PASS） |
| `rpfem_native_x64.dll` | DLL（Fix9 と同じもの、SHA256 `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`。rc4b では DLL を変えていません） |
| `WORKBOOK_VALIDATION.json` | コンパイル・開き直しの確認記録 |

ブックと DLL は同じフォルダに置いてください。ハッシュは `SHA256SUMS.txt` にあります。

## 確かめたこと

- Winmax2 で 7 ケース（P10A、D05A、P02A、P05A、P04A、D07A、強制薄層の P02A_THIN）を確認中です。同じ仕組みの rc4（θ 0.9）では、6 ケースは Fix9 と Fs・witness が一致し、P02A_THIN は GAP 0.450%（目標 0.5%）でした。rc4b の結果は PR に追記します。
- **P06N・P06A は VBA ではまだ確かめていません。** Python の再現（VBA と同じ 400 要素の粗いメッシュから開始）だけです：P06N GAP 1.69%、P06A 1.52%（2 台で同じ結果）。VBA の時間の見積りは各 2 時間弱ですが、測った値ではありません。

## 元に戻すには

- `ADAPT_THIN_EARLY_UPPER_THETA=0.9` で rc4 と同じ、さらに `ADAPT_THIN_EARLY_UPPER_THETA=0.8`・`ADAPT_THIN_EARLY_UPPER_CAP=3500` で rc3b と同じ数値になります。
- `ADAPT_THIN_EARLY_START=0` で早期開始を止めます。

`tools/` はビルド用です（Fix9rc2 の版フォルダを前提にしています）。差分は `docs/` にあります。
