# G1R20 Fix8（2026-10-19）

G1R20 Fix7（main `a12c8b9`）をもとにした版です。主な内容は次の2つです。

1. **P06（薄い材料層）向けのメッシュ手法の移植**：Python 試作で確かめた手法を VBA に入れました。ただし、**P06N/P06A ではまだ検証していません**（下記）。
2. **精度ラウンドを行うかどうかの判定を、時間ではなく要素数で決める**ようにしました（Issue [#29](https://github.com/tak063495-prog/excel-vba-rpfem/issues/29)）。PC の速さによって計算の経路が変わらないようにするためです。

方針はこれまでと同じで、**精度確保が先、速度改善はその次** です。

- 配布ブック：`delivery/RPFEM_20261019_G1R20_Fix8.xlsm`（SHA256 `614279ef1f952622facb0bc54c97757599f6c32a782e429317659d8aa99acb82`）
- ビルド検証：`delivery/WORKBOOK_VALIDATION.json`（state PASS、VBA コンパイル PASS、保存後の再オープン＋再コンパイル PASS、ブック内 63 モジュールとソースの SHA256 一致）
- DLL は Fix4〜Fix7 と同じです（`rpfem_native_x64.dll`、SHA256 `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`）
- 保存時の設定は Fix5〜Fix7 と同じ 2% 標準プリセット（`RPX_Speed2pctPreset`）です
- ソースは候補 rc2（コミット `511cc5b`）です。変更したモジュールは `RPX_AccuracyNative.bas`、`RPX_Adapt.bas`、`RPX_Refine.bas`、`RPX_Solve.bas` の4つです（`SOURCE_MANIFEST.json` の `changed_modules`）
- ファイルの SHA256 は `SHA256SUMS.txt` にあります

## 精度：変わっていません

Winmax2（AC 電源、電源設定は変更なし）で、Fix7 rc2（配布版）と同じハーネスで測りました。

| ケース | 状態 | Fs（下界〜上界） | GAP | 要素数（下界/上界） | 比較基準との一致 |
|---|---|---|---:|---|---|
| P10A | PASS_TARGET | 2.00606〜2.04459 | 1.902% | 4018 / 790 | Fs 全桁一致、witness 一致 |
| D05A | PASS_TARGET | 1.30142〜1.32029 | 1.439% | 355 / 355 | 同上 |
| P02A | PASS_TARGET | 1.30129〜1.32149 | 1.540% | 355 / 355 | 同上 |
| P05A | PASS_TARGET | 1.46421〜1.49342 | 1.975% | 656 / 656 | 同上 |
| P04A | PASS_TARGET | 2.24914〜2.28475 | 1.571% | 355 / 355 | 同上（同じ日の Fix7 rc2 対照） |
| D07A | PASS_TARGET | 1.26427〜1.28091 | 1.308% | 355 / 355 | 同上（同じ日の Fix7 rc2 対照） |
| P02A_THIN（強制、目標 0.5%） | PASS_TARGET | 1.30298〜1.30923 | 0.479% | 1033 / 707 | Fix8 rc1 と一致 |

**全ケースで Fs は Fix7 と全桁同じで、`witness_lower.bin`/`witness_upper.bin` もバイト単位で同じです。精度の低下はありません。**
これらのケースでは、薄い層の判定は働いていません（P10A・P05A のログ `thin_layer_detect active=0`）。下界の直接修正（`accuracy_dll_equality_direct`）も一度も動いていません。

## 時間：速度改善とは確定していません

| ケース | Fix8 rc2 | 比較基準 | 
|---|---:|---|
| P10A | 1340.5 秒 | 同じ日の Fix7 rc2 再実行 1740.3 秒、Fix8 rc1 再実行 1732.0 秒 |
| D05A | 163.0 秒 | Fix7 rc2 242.2 秒、Fix8 rc1 241.9 秒 |
| P02A | 199.0 秒 | Fix7 rc2 295.1 秒、Fix8 rc1 296.4 秒 |
| P05A | 967.8 秒 | Fix7 rc2 1465.1 秒、Fix8 rc1 1476.5 秒 |
| P04A | 192.8 秒 | 同じ日の Fix7 rc2 対照 283.5 秒 |
| D07A | 187.2 秒 | 同じ日の Fix7 rc2 対照 275.4 秒 |

- P10A では、rc2 の要素数の判定で 1109 要素のラウンドを行わなくなりました。このラウンドはこれまで P10A で行われた4回すべてで HSD の時間上限で打ち切られ、結果に使われていません（1回あたり約 380 秒）。P10A の短縮の多くはこの分と見ていますが、**これは「候補」です。**
- 理由：同じ rc2 のバッチで、rc2 の変更が動いていないケース（D05A、P02A、P05A、P04A、D07A）も 32〜34% 速く測れました。求解回数とログの行数は Fix7 と同じで、すべての段階が同じ割合で速くなっています。測定時の PC の状態の違いと見ています（原因は未特定）。P10A も同じ影響を受けている可能性があります。
- **確定するには、Fix7 rc2 と Fix8 rc2 を同じ条件で交互に測る必要があります（未実施）。**

## 変更内容

### 1. 精度ラウンドの判定を要素数で行う（rc2、`RPX_Adapt.bas`、`RPX_Solve.bas`）

Fix7 では、次の精度ラウンドの上界の時間を「直前の HSD 時間 × 要素数の比」で見積もり、450 秒を超えるとそのラウンドを行いませんでした。
Fix8 rc1 の P10A では、PC が一時的に約 1.45 倍遅くなっていたため、790 要素のラウンドの見積もりが 488.7 秒となって止まり、**GAP 2.066%（未達）** になりました。計算の中身は Fix7 と全桁同じで、違いは時間だけでした（同じ条件の再実行では rc1 も Fix7 も 1.902%）。

rc2 では次のようにしました。

- `ADAPT_ACCURACY_SKIP_ESTIMATE=2`（既定）：候補が `ADAPT_ACCURACY_SKIP_ELEMENTS`（既定 1000）要素を超えるラウンドは行いません。
- 行うラウンドでは、HSD の時間上限を `ADAPT_ACCURACY_HSD_SECONDS`（既定 900 秒）まで引き上げます。下げることはありません。遅い PC で打ち切られないための安全側の上限です（ログ `accuracy_hsd_floor`）。
- 時間は、全体の時間予算による安全停止にだけ使います。
- 適用範囲は Fix7 と同じで、上界 conic の省略が有効なときだけです。`ADAPT_ACCURACY_SKIP_ESTIMATE=1` にすると Fix7 の時間見積もりに戻ります。
- 1000 要素という境目は、Winmax2 のこれまでの記録から決めました。この判定が使われたのは P10A（582/790/1109 要素）だけでした。582 要素と 790 要素のラウンドは行い、これまで成功したことのない 1109 要素のラウンドは行いません。

注意（表示だけの問題）：要素数の判定でラウンドを行わなかったときも、`accuracy_stop` の理由は Fix7 と同じ `upper_solve_estimate_exceeds_limit` と記録されます。判定の内容は、その直前の `accuracy_upper_estimate mode=elements;elements=1109;max_elements=1000` の行で確認できます。表示を直すとブックが変わり再検証が必要になるため、Fix8 では直していません（次の版で直す予定）。

### 2. 薄い材料層のメッシュ手法（rc1、`RPX_Refine.bas`、`RPX_Adapt.bas`）

P06 の Python 試作で確かめた手法です。**薄い材料層があるモデルだけで動きます。**

- 判定（`ADAPT_THIN_LAYER`：0＝使わない、1＝自動（既定）、2＝必ず使う）：材料ごとの薄さ（2×面積 ÷（周長/2）²）が `ADAPT_THIN_LAYER_RATIO`（既定 0.1）より小さい材料があるときだけ使います。これまでのモデルで該当するのは P06 の帯の層（0.027）だけです。ほかは 0.18 以上です。Fs モード、水なし、剛体なし、材料ごとに単位体積重量が一定、のときに限ります（ログ `thin_layer_detect`）。
- 使うときは、通常の適応計算の後に、**上界と下界それぞれ専用のメッシュの系列** を作ります（上界を先に1段、次に下界を1段、を交互に繰り返す）。
  - 上界：誤差指標（ひずみ速度と速度の跳び）で辺を選び、要素の向きに合わせて分割（方向分割）し、最初の段では帯の中央線も入れます。その後、辺の入れ替え（`RPX_FlowMesh`、5回）でメッシュを整えます。
  - 下界：上界の速度場の情報と下界の応力場の情報を混ぜた指標で要素を選んで分割し、扇形の細分（`RPX_AngularFanMesh`）を入れます。
  - 中央線は、薄さが `ADAPT_THIN_LAYER_MID_RATIO`（既定 0.25）より小さい材料すべてに入れます。帯の層だけに入れると、Python 試作で下界が悪くなったためです（P06N：0.6668／2964 要素 対 0.6767／3528 要素）。
  - 各段の Fs は、本計算と同じ根の規則（q8、Fs 許容差、`SaveState`＋監査、証明書）で求めます。良くなったときだけ結果を置き換えます。
  - 主な設定と既定値：上界 10 段（`ADAPT_THIN_LAYER_UPPER_ROUNDS`）、下界 3 段（`_LOWER_ROUNDS`、2 段目から開始 `_LOWER_START`）、選ぶ割合 0.8／0.9（`_UPPER_THETA`／`_LOWER_THETA`）、要素数の上限 3500／4000（`_UPPER_CAP`／`_LOWER_CAP`）。
  - ログ：`thin_layer_begin`、`thin_layer_step`、`thin_layer_rejected`、`thin_layer_end`。

### 3. 下界の等式の直接修正（rc1、`RPX_AccuracyNative.bas`）

`RecoverEquilibrium` の LSMR による修正がうまくいかず、残差が 1e-5 以下のときだけ、DLL の直接法（LDL）で応力だけを最小限に動かして等式を満たします（荷重は動かしません）。結果は、これまでと同じ完全な監査（`RPX_AccuracyPrimalAudit`、`RPX_FieldGate`）に通ったときだけ使います（設定 `ACCURACY_DLL_EQUALITY_DIRECT`、既定 1、ログ `accuracy_dll_equality_direct`）。
今回の 7 ケースでは一度も動いていません。

## P06N/P06A について（未検証）

- 薄い材料層の処理は、**P02A に `ADAPT_THIN_LAYER=2`（強制）を設定したテスト（P02A_THIN）でだけ動作を確かめました**。上界 1 段（707 要素、辺の入れ替え 176 回）、下界 1 段（1033 要素）がどちらも監査に通り、GAP 0.479% になりました。
- **P06N/P06A では VBA で一度も実行していません。** P06 でこの処理が働くか、時間内に GAP 2% に届くかは未確認です（別PCで検証予定）。P06 の帯の層では自動で働く設計です。

## 検証データ

| フォルダ | 内容 |
|---|---|
| `validation/G1R20_Fix8_rc2_optest_20261019` | **この配布版（rc2、`511cc5b`）** の 7 ケース |
| `validation/G1R20_Fix8_optest_20261019` | rc1（`42c6cc0`）の 7 ケース。P10A が PC の遅い状態で 2.066%（未達） |
| `validation/G1R20_Fix8_p10rerun_20261019` | rc1 の P10A 再実行（1.902%、1732.0 秒） |
| `validation/G1R20_Fix8ctl_p10rerun_20261019` | 直後に行った Fix7 rc2 の P10A 対照（1.902%、1740.3 秒） |
| `validation/G1R20_Fix8ctl_fix7_20261019` | Fix7 rc2 の P04A/D07A 対照 |

それぞれ results/*.json、cases/*/analysis.tsv、witness などを含みます。大きな作業用ブック（`workbooks/`、`prepared_template.xlsm`）は含めていません。

## 今後の候補（未検証）

- Fix7 rc2 と Fix8 rc2 を同じ条件で交互に測り、P10A の短縮を確かめること
- `accuracy_stop` の理由の表示を直すこと（要素数の判定のとき）
- 1000 要素の境目をほかのモデルで確かめること
- P06N/P06A での検証（別PC）
