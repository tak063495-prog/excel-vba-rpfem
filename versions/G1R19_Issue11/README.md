# G1R19 Issue #11 修正版（適格な q8 保存場の pre-final 再利用が load phase 不一致で拒否される）

[Issue #11 のコメント](https://github.com/tak063495-prog/excel-vba-rpfem/issues/11#issuecomment-6204373784) の指摘への修正です。基準は main `46a27db`（G1R19）。G1R19 フォルダーは変更していません。

## 利用

- `delivery/RPFEM_20261016_G1R19_Issue11.xlsm` と同じフォルダーに `rpfem_native_x64.dll`（G1R19 と同一、SHA256 `66da1a21…`）を置きます。ブックは G1R19 配布ブックの標準モジュール `RPX_Adapt` だけを置換したものです。
- `src/vba/RPX_Adapt.bas` が編集元、`import/RPX_Adapt.bas` が取込用（同一内容）、差分は `changes.diff`。

## 原因

`AfterCycles` は pre-final 再利用判定（`RPX_TryTerminalPairReuse`）の前に `RPX_RestoreState(RBestPair)` を呼び、`RPX_ReadInputs`/`RPX_AssignRegions` 経由で `RPX_LoadPhase="raw"` に戻します。一方、保存場は Fs 評価で `RPX_TotalReferenceLoads` を通った `reference_total` の状態です。`TerminalOwnMetadataReady` の `state.phase <> RPX_LoadPhase` と、続く `RCurrent.originalModel <> RPX_ExactModel()` が不一致になり、再利用条件を満たす q8 pair でも必ず `Evaluate(mode,"final")` を再実行していました。

## 修正（`RPX_TryTerminalPairReuse` のみ）

- phase／exact-model の照合は削除・緩和していません。照合の前に、保存 pair 自身の荷重・材料・境界・メッシュ文脈を `RPX_RestoreWitnessState(RCurrent)` で復元します（入力世代・方針・exact-model の一致はこの関数が検査し、不一致なら例外→不成立）。そのうえで従来どおり exact-model、上下界それぞれのメタデータ（q8・証明書・根区間幅・世代・方針・phase）、元メッシュでの場の再監査を行います。
- 不成立時は `RPX_RestoreState(baseline)` で、呼び出し前と同じ raw 入力状態に戻し、既存の `Evaluate("final")` に進みます。成立時は保存場の `reference_total` 文脈で続行します（`Evaluate("final")` 後と同じ状態で、出力側の exact-model 検査も通ります）。
- 呼び出し時点で既に同じ phase の場合（final 後の2回目の呼び出し）は G1R19 と同じ処理順・復元方法のままです。

## 検証（2026-10-16、Rico 実Excel 16.0、自分で起動した Excel のみ）

### 制御経路（`tests/run_issue11_excel.py`、結果 `results/issue11_excel.json`）

小規模 fixture（`tests/Issue11Probe.bas`、試験専用 GAP=1.9、CYCLES=1 で `AfterCycles` の `RestoreState` を通す）。

|ケース|G1R19|修正版|
|---|---|---|
|最初から q8 の適格 pair（`ADAPT_PRECISION_FROM_START=1`）|final 評価 1 回、Fs 呼び出し 4|**final 評価 0 回**、Fs 呼び出し 2、上下界はビット一致（0.615250856103897／0.728120003688654）|
|q4 粗探索の pair（不適格、`=0`）|final 評価 1 回、4 呼び出し|final 評価 1 回、4 呼び出し（必須の q8 再評価は従来どおり）|
|`TerminalReuseSavedPairProbe`（q/世代/方針/根区間/場破損など16種の不正で拒否）|PASS|PASS|

配布ブックは置換・コンパイル・保存・再オープン後に `RPX_Adapt` がソースと一致し、他モジュールは大文字小文字の表記差のみ（PASS）。

### 実ケース（標準設定 GAP=0.01, HSD_SOLVE_SECONDS=300, DIRECT_DLL=1, native auto）

G1R19 で目標未達・低速・精度低下だった 10 ケースを再実行。全ケースで Fs 上下界・gap は G1R19 と一致し悪化なし、監査はすべて PASS。これらのケースの終端 pair は q4 粗探索由来で再利用対象外のため、本修正の経路は通りません（G1R17 比の低速化・精度低下は残存）。詳細は `results/native_cases.md`、原データは `results/comparison.json`。
