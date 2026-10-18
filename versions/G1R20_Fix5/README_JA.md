# G1R20 Fix5（2026-10-18）

G1R20 Fix4 をもとに、Issue [#28](https://github.com/tak063495-prog/excel-vba-rpfem/issues/28)（D05A の停止）と Issue [#29](https://github.com/tak063495-prog/excel-vba-rpfem/issues/29)（P06N/P06A の GAP 未達・時間の無駄）に対応した版です。
方針は **精度確保が先、速度改善はその次** です。同じマシンで「速くなり、しかも GAP 2% 以内を保つ」と測定で確かめた設定だけを「速度改善」と呼びます。確かめていないものは「候補」として分けて書きます。

DLL は Fix4 と同じです（`rpfem_native_x64.dll`、SHA256 `66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`）。

## 修正内容

### #28：精度DLLの追加回復が不採用でも解析を止めない（`RPX_Fs.bas`、`RPX_Guard.bas`）
- Fix4 では、HSD が OPTIMAL で原物理監査にも合格した後、精度DLLの原dual回復（`RPX_AccuracyRecoverDual`）や下界の原dualゲート（`RPX_OriginalLowerDualGate`）が不採用になると、`HSD_NOT_CONVERGED` として Fs 探索全体が止まっていました（D05A：11秒で失敗）。
- Fix5 では、不採用を `ACCURACY_DUAL_RECOVERY_REJECTED`／`ACCURACY_LOWER_DUAL_GATE_REJECTED` という実態どおりの名前で記録します。既存の回復経路（原dual・primal witness）は従来どおり先に試し、すべて不採用のときだけ、**RPX_Bound で原物理監査に合格した HSD OPTIMAL の点** を使って探索を続けます（`ACCURACY_DLL_RECOVERY=0` と同じ根拠）。ログには `accuracy_recovery_fallback` が残ります。
- HSD 自体が未収束・時間切れの点（P06 の場合）にはこの代替を使いません。原物理監査に合格していない点も使いません。
- 下界側の代替は、ゲートが精度DLLのために有効な場合だけです（`LOWER_FIXED_RECOVERY=1` を明示した場合は従来どおり）。
- 原dual監査の stationarity 閾値は **緩めていません**。[PR #30](https://github.com/tak063495-prog/excel-vba-rpfem/pull/30) の検証で、相対閾値（1e-7×‖cost‖∞）にしても raw 残差 1429 は約1000万倍超過で救えず、閾値を緩めるだけでは誤判定の危険が増えると分かったためです。正規化 cost での正しい復元は今後の候補です。

### #29：同じメッシュ・同じ q8 の解き直しを省く（`RPX_Adapt.bas`）
- Fix4 は GAP 2% 未達のとき、最後の独立 q8 確認（own_q8）で、すでに q8・同じメッシュで根が確定している結果を最初から解き直していました（P06N で約4370秒、P06A で約1170秒、Fs は完全に同じ）。
- Fix5 は辺ごとに、この解析で認定済みの「自メッシュ・q8・根幅が許容内」の結果なら、原物理場を再監査するだけで採用し、解き直しません（`own_quadrature_reuse`）。条件は既存の `RPX_TryTerminalPairReuse` と同じ照合を片側ずつ行うもので、GAP 未達でも使えます。照合に通らなければ従来どおり解き直します。

### #29：候補（既定は無効。プリセットにも入れていません）
- `ADAPT_OPTIONAL_DEADLINE_SECONDS`（既定0＝無効）：解析開始からこの秒数を過ぎたら、任意の追加段（精度ラウンド、下界 fan、C0 wedge）を始めずに、確定済みの上下界を出力します。外側の時間上限で結果が失われる問題（#29 の HSD 900秒試験）への対策候補です。
- `HSD_SOLVE_SECONDS_PER_1000_ELEMENTS`（既定0＝無効）：HSD 1回の時間上限を「要素数/1000×この値」まで引き上げます（下げない、上限3600秒）。#29 では一律900秒にしても改善しなかったため、効果は未確認です。
- どちらも P06N/P06A での検証が済むまで「候補」です。

### 変更しなかった点と理由
- **上界メッシュの細分方針**：Python の 2% 対は既知形状から別経路（辺の組替え・局所／方向細分で 8739 要素）で得たもので、通常解析の細分に移せる具体的な変更点は、今回の資料からは特定できませんでした。P06 の検証結果を見てから見直します。
- **`ADAPT_PRECISION_FROM_START`**：2%標準プリセットでは 1 のままです（G1R19 速度設定は 0）。Fix4 の 26 ケースでは 1 で 21 ケースが目標達成しており、0 との同一マシン比較はまだありません。q8 を最初から使うことで上記の own_q8 再利用が効く条件にもなっています。候補として P06 の検証で比較します。
- **「収束していない計算を Fs 探索の手がかりに使わない」**：ログを見直すと、#29 の 0.5591 は双対ガイドではなく、前段の下界 0.6213 に探索幅 10% を掛けた初期点（0.6213×0.9）である可能性が高いと判断しました（推定）。未収束点の扱いは既存コードで「符号は独立監査済みの証明だけ・ガイドがある区間は二分法」となっており、今回は変更していません。

### 画面表示
- 設定シートのボタン名を「2%標準の設定（速度効果は未検証）」「G1R20・改善候補（未検証）を選択」に変えました。マクロ名と設定値は Fix4 と同じです。

## プリセットの意味と検証状況

| プリセット（マクロ） | 意味 | 検証状況 |
|---|---|---|
| 2%標準の設定（`RPX_Speed2pctPreset`） | GAP 2%、q8、直接DLL（`DIRECT_DLL=1`）、精度DLL回復あり、最初から高精度（`ADAPT_PRECISION_FROM_START=1`）、HSD 300秒、精度ラウンド上限8・24000要素 | Fix4 で 26 ケース中 21 ケースが GAP 2% 以内（Triela）。Fix5 は下記4ケースで動作確認。**速度改善としては未検証**（旧名「速度優先」は歴史的な名前です） |
| G1R20・改善候補（`RPX_G20ImprovementPreset`） | 上の設定に、balanced 下界候補、上界 flow 細分5回、終端スケジュール、再中心化、dual線形補正などを追加 | **候補**。Fix4 の P06N では各段 1800〜2800 秒遅く、上下界はほぼ同じ（#29）。改善は確認されていません |
| 通常（Setup 直後） | `DIRECT_DLL=0`、`ADAPT_PRECISION_FROM_START=0`、GAP 2% | 従来どおり |

配布ブックは「2%標準の設定」を選んだ状態で保存しています（Fix4 は改善候補を選んだ状態でした）。

## 導入
`delivery` の XLSM と `rpfem_native_x64.dll` を同じフォルダーに置きます（64bit Excel）。既存ブックの更新は Fix4 と同じ手順で、63 標準モジュールを `src/vba`（UTF-8）または `import`（CP932）で置き換えます。Fix4 からの変更は `RPX_Adapt.bas`、`RPX_Fs.bas`、`RPX_Guard.bas`、`RPX_Solve.bas`、`RPX_UI.bas` の5つです（`SOURCE_MANIFEST.json` の `changed_modules`）。

## ビルド
Winmax2（Ryzen 7 7840U、Excel 365 64bit 16.0.20430）で `tools/build_candidate_book.py` を実行し、Fix4 配布ブックへ全63モジュールを読み込み、コンパイル・保存・再オープン・再コンパイル・プリセット値の照合を行いました（`delivery/WORKBOOK_VALIDATION.json`、state=PASS）。

配布ファイル（`delivery/`）の SHA256：
- `RPFEM_20261018_G1R20_Fix5.xlsm`：`601985a2fcac3d7ad0040b6ca972f9b5ad20454b1da7ee753e16b14c812dd66f`
- `rpfem_native_x64.dll`：`66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b`（Fix4 と同一）
- `WORKBOOK_VALIDATION.json`：`2f6ece9afe4717c770061117a28627de39587fe8c3604fa822cfecc90af43f97`

## 動作テスト結果（2026-10-18、Winmax2）

条件：Winmax2（GPD G1619-04、Ryzen 7 7840U 8コア16スレッド、RAM 48GB、Windows 11、Excel 365 64bit）、DLL 使用（`DIRECT_DLL=1`）、2%標準プリセット（`RPX_Speed2pctPreset`）、1ケースずつ逐次実行。対象はユーザー指定の「Fix4 で成績が悪かった短時間ケース」と D05A の相似元 P02A です。**P06N/P06A はユーザーが別の PC で検証するため、ここでは実行していません。**
記録：[`validation/G1R20_Fix5_optest_20261018`](../../validation/G1R20_Fix5_optest_20261018/)、Fix4 対照：[`validation/G1R20_Fix4_control_P10A_20261018`](../../validation/G1R20_Fix4_control_P10A_20261018/)。各ケースの最終上下界は独立監査（Python の原物理場監査）に合格しています。

| ケース | Fix5（Winmax2） | Fix4（Triela、26ケース計測） | 判定 |
|---|---|---|---|
| D05A | **PASS_TARGET** 1.30142〜1.32029、GAP 1.44%、187.8秒、355/355要素 | FAILED（`HSD_NOT_CONVERGED`、11.4秒） | #28 解消 |
| P02A（D05Aの相似元） | PASS_TARGET 1.30129〜1.32149、GAP 1.54%、226.6秒 | PASS_TARGET 同値、234.4秒 | 同値 |
| P05A | **PASS_TARGET** 1.46421〜1.49342、GAP 1.98%、1166.6秒 | TIMEOUT_STOPPED（上限1200秒、Fs は同値） | Fs 同値、上限内で完了 |
| P10A | FS_AVAILABLE_PARTIAL 2.00606〜2.04794、GAP 2.07%、1686.3秒 | FS_AVAILABLE_PARTIAL 同値、2261.7秒 | 同値（2% 未達は Fix4 と同じ） |

- D05A は、Fix4 で `ACCURACY_DLL_RECOVERY=0` にした診断（1.30142〜1.32029、GAP 1.44%）と同じ値です。ログでは `accuracy_recovery_fallback` が4回記録されています。
- P10A では最後の q8 確認で `own_quadrature_reuse` が上下両側に働き、解き直しはありませんでした。**同じ Winmax2 で Fix4 を回した対照**は、同じ Fs・GAP で 2080.4秒、そのうち q8 の解き直しが約367秒（下界80秒＋上界287秒）でした。Fix5 は 1686.3秒です。ただし 1ケース・1回だけの比較なので、「速度改善」とはまだ扱わず候補とします。
- Triela と Winmax2 は別のマシンなので、上の表の時間を比べて速度の優劣は主張しません。
- **P10A の1回目について**：1回目（`cases/P10A_run1_on_battery`）は、ノートPCがバッテリー駆動（バランス電源）だった時間帯に実行され、最初の求解（coarse_0、374要素）で HSD が1反復約9秒かかり、300秒の上限に当たって `HSD_SOLVE_BUDGET_EXCEEDED` で停止しました。AC 電源に戻ったあとの Fix4 対照と Fix5 の2回目では、同じ求解が約4秒/反復・約180〜190秒で OPTIMAL になっています（目的値も一致）。この求解の経路は Fix5 で変更していません。HSD の300秒上限が遅いマシンでは足りない場合がある点は、今後の注意点です（`HSD_SOLVE_SECONDS_PER_1000_ELEMENTS` などの候補で検討）。

## 未検証・残りの作業
- **P06N/P06A**：ユーザーが別の PC で Fix5 を使って検証する予定です。#29 は開いたままにします。
- 26ケース全体の Fix5 での再計測はまだです。
- 上記の候補設定（`ADAPT_OPTIONAL_DEADLINE_SECONDS`、`HSD_SOLVE_SECONDS_PER_1000_ELEMENTS`、`ADAPT_PRECISION_FROM_START` の 0/1 比較、改善候補プリセット）は、同じマシンでの測定で精度を保ったまま速いと確かめるまで「候補」のままです。
