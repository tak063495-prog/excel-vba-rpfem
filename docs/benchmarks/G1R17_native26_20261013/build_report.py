"""Regenerate the published report and figures from recorded data; never run Excel.

Python 3.12 + matplotlib + numpy. The default command uses data/summary.json.
--source imports the immutable native26 measurement directory once for publication.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import re
import shutil
from pathlib import Path

import matplotlib

matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib import font_manager
from matplotlib.lines import Line2D
from matplotlib.patches import Patch

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[2]
ZIP_NAME = "RPFEM_G1R17_Native26_20261012.zip"
ZIP_SHA = "3b65c2f3c2ed20283ae9f052321785a7547988391895cfea37e475082e4f6d62"
BASE_URL = "https://github.com/tak063495-prog/excel-vba-rpfem"
TEAL, AMBER, INK = "#087f82", "#c87b13", "#25374b"


def read_json(path):
    return json.loads(Path(path).read_text(encoding="utf-8-sig"))


def write_json(path, value):
    Path(path).write_text(json.dumps(value, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")


def sha256(path):
    with Path(path).open("rb") as stream:
        return hashlib.file_digest(stream, "sha256").hexdigest()


def import_records(source):
    data = HERE / "data"
    data.mkdir(parents=True, exist_ok=True)
    (data / "inputs").mkdir(exist_ok=True)
    original = read_json(source / "results/final_summary.json")
    metadata = read_json(source / "results/environment.json")
    cases = []
    selected = (
        "case", "status", "analysis_seconds", "lower_fs", "upper_fs", "relative_gap",
        "fs_obtained", "target_met", "target_unmet_reasons", "element_counts", "witness_q",
        "error_number", "pore_pressure", "root_width", "root_search_incomplete",
        "global_incomplete_note", "dll_actually_used", "native_counters", "input_sha256",
        "source_sha256", "dll_sha256", "workbook_sha256", "production_source_check",
        "independent_audit", "actual_settings", "root_tolerance_contract",
    )
    for c in original["cases"]:
        record = {key: c[key] for key in selected}
        input_path = source / "inputs" / (c["case"] + ".json")
        record["input_title"] = read_json(input_path)["source_case"]["title"]
        record["solve_count"] = c["log_summary"]["solve_count"]
        record["audit_count"] = c["log_summary"]["audit_count"]
        record["native_terminal"] = c["log_summary"]["terminal"]
        cases.append(record)
        shutil.copy2(input_path, data / "inputs" / input_path.name)
    summary = {k: v for k, v in original.items() if k != "cases"}
    summary.update(metadata=metadata, cases=sorted(cases, key=lambda c: c["case"]))
    write_json(data / "summary.json", summary)
    for name in ("RESULTS.csv", "PHASE_TIMINGS.csv"):
        shutil.copy2(source / name, data / name)
    for name in ("quality_notes.json", "prior_dll2_comparison.json", "final_verification.json", "environment.json"):
        shutil.copy2(source / "results" / name, data / name)
    shutil.copy2(source / "QUALITY_NOTES.md", HERE / "QUALITY_NOTES.md")


def configure_plotting(font_path):
    if font_path:
        font_manager.fontManager.addfont(font_path)
        family = font_manager.FontProperties(fname=font_path).get_name()
    else:
        candidates = ["Yu Gothic", "Noto Sans CJK JP", "IPAexGothic", "MS Gothic"]
        installed = {font.name for font in font_manager.fontManager.ttflist}
        family = next((x for x in candidates if x in installed), None)
        if family is None:
            raise RuntimeError("Japanese font required; supply --font with a local Japanese font file.")
    plt.rcParams.update({
        "font.family": family, "font.size": 12, "axes.labelsize": 12,
        "axes.titleweight": "bold", "axes.unicode_minus": False,
        "text.color": INK, "axes.labelcolor": INK, "xtick.color": INK,
        "ytick.color": INK, "axes.edgecolor": "#c9d2db", "svg.fonttype": "path",
        "savefig.facecolor": "white", "figure.facecolor": "white",
    })


def decorate(ax, labels):
    ax.set_yticks(np.arange(len(labels)), labels)
    ax.invert_yaxis()
    ax.set_axisbelow(True)
    ax.grid(axis="x", color="#e3e8ed", linewidth=0.8)
    ax.tick_params(axis="both", length=0)
    for side in ("top", "right", "left"):
        ax.spines[side].set_visible(False)


def status_legend(ax, lower=False):
    handles = [Patch(color=TEAL, label="目標達成：10ケース"), Patch(color=AMBER, label="Fs取得・目標未達：16ケース")]
    if lower:
        handles += [Line2D([], [], color=INK, marker="|", markersize=12, linestyle="", label="下界 / 上界の端点")]
    ax.legend(handles=handles, loc="lower left", bbox_to_anchor=(0, 1.005), ncol=2, frameon=False, fontsize=11)


def save(fig, name):
    folder = HERE / "figures"
    folder.mkdir(exist_ok=True)
    fig.savefig(folder / (name + ".png"), dpi=180)
    fig.savefig(folder / (name + ".svg"))
    plt.close(fig)


def make_figures(summary, previous):
    cases = summary["cases"]
    rows = sorted(cases, key=lambda c: -c["analysis_seconds"])
    fig, ax = plt.subplots(figsize=(12, 11))
    fig.subplots_adjust(left=0.13, right=0.91, bottom=0.09, top=0.84)
    y = np.arange(len(rows))
    minutes = [c["analysis_seconds"] / 60 for c in rows]
    ax.barh(y, minutes, height=0.67, color=[TEAL if c["target_met"] else AMBER for c in rows])
    decorate(ax, [c["case"] + (" 水圧" if c["pore_pressure"] else "") for c in rows])
    ax.set_xlim(0, 112)
    ax.set_xlabel("主解析時間（分）  /  run_end.elapsed_s")
    for i, value in enumerate(minutes):
        ax.text(value + 0.8, i, f"{value:.2f}", va="center", fontsize=11)
    status_legend(ax)
    fig.text(0.07, 0.957, "G1R17 最新DLL5：26ケースの実測時間", fontsize=22, weight="bold")
    fig.text(0.07, 0.917, "26/26でFs上下界を取得  ｜  合計 7時間46分7秒  ｜  各ケース1回・順次実行", fontsize=12)
    fig.text(0.07, 0.025, "初期メッシュ・水圧準備・Excel起動終了・保存・Python監査の時間は除外。2026-10-12〜13測定。", fontsize=10.5)
    save(fig, "runtime_26")

    fig, ax = plt.subplots(figsize=(12, 11))
    fig.subplots_adjust(left=0.13, right=0.69, bottom=0.1, top=0.83)
    y = np.arange(len(cases))
    for i, c in enumerate(cases):
        color = TEAL if c["target_met"] else AMBER
        ax.plot([c["lower_fs"], c["upper_fs"]], [i, i], color=color, lw=3,
                marker="|", markersize=11, markeredgewidth=1.8)
        ax.text(1.035, i, f'{c["lower_fs"]:.6f} – {c["upper_fs"]:.6f}',
                transform=ax.get_yaxis_transform(), va="center", fontsize=10.5, color=color)
    decorate(ax, [c["case"] for c in cases])
    ax.set_xlim(0.45, 2.38)
    ax.set_xlabel("Fs（下界から上界まで）")
    ax.axvline(1, color="#687c92", ls="--", lw=1.1)
    ax.text(1.035, 1.028, "Fs下界 – Fs上界", transform=ax.transAxes, fontsize=11, weight="bold")
    status_legend(ax)
    fig.text(0.07, 0.957, "26ケースのFs上下界：全て独立物理監査を通過", fontsize=21, weight="bold")
    fig.text(0.07, 0.917, "A：関連流れ則  /  N：現行Davis等価関連モデル  /  破線：Fs = 1", fontsize=12)
    fig.text(0.07, 0.032, "線分は計算上の上下界区間。Nは非関連流れ則そのものの厳密上下界を示すものではない。", fontsize=10.5)
    fig.text(0.07, 0.01, "採用積分次数は表を参照（P08A/P09A/P10Aは4/4、D02Aは4/8、D07Aは8/4）。", fontsize=10.5)
    save(fig, "fs_intervals_26")

    rows = sorted(cases, key=lambda c: -c["relative_gap"])
    fig, ax = plt.subplots(figsize=(12, 11))
    fig.subplots_adjust(left=0.13, right=0.91, bottom=0.1, top=0.84)
    gaps = [100 * c["relative_gap"] for c in rows]
    bars = ax.barh(np.arange(len(rows)), gaps, height=0.67,
                   color=[TEAL if c["target_met"] else AMBER for c in rows])
    for bar, c in zip(bars, rows):
        if c["case"] in {"D07A", "P08A", "P09A", "P10A"}:
            bar.set_hatch("///")
            bar.set_edgecolor("#704308")
    decorate(ax, [c["case"] for c in rows])
    ax.axvline(1, color=INK, ls="--", lw=1.1)
    ax.set_xlim(0, 13.8)
    ax.set_xlabel("上下界GAP（%）  =  100 × (上界 − 下界) / 平均値")
    for i, value in enumerate(gaps):
        ax.text(max(value + 0.1, 1.16), i, f"{value:.3f}", va="center", fontsize=11)
    status_legend(ax)
    fig.text(0.07, 0.957, "精度目標の達成は10/26：上下界幅と最終確認", fontsize=21, weight="bold")
    fig.text(0.07, 0.917, "破線はGAP目標1%。最終確認・根探索・独立監査も満たした場合に目標達成と判定。", fontsize=12)
    fig.text(0.07, 0.032, "斜線：最終確認未完了。P08A/P09A/P10Aはq8のHSD時間予算に達し、監査済み旧場を保持。", fontsize=10.5)
    fig.text(0.07, 0.01, "D07AはGAPが1%以下でも完了フラグ未達。P06A/P06Nはq8上下界を取得し、GAPが目標未達。", fontsize=10.5)
    save(fig, "gap_26")

    current = {c["case"]: c for c in cases}
    old_rows = sorted(previous["cases"], key=lambda c: c["case"])
    old_total = sum(c["seconds"] for c in old_rows)
    new_total = sum(current[c["case"]]["analysis_seconds"] for c in old_rows)
    fig, ax = plt.subplots(figsize=(12, 7.8))
    fig.subplots_adjust(left=0.13, right=0.9, bottom=0.15, top=0.76)
    y = np.arange(len(old_rows))
    ax.barh(y - 0.18, [c["seconds"] / 60 for c in old_rows], height=0.32, color="#8fa2b9", label="過去DLL2実測")
    ax.barh(y + 0.18, [current[c["case"]]["analysis_seconds"] / 60 for c in old_rows], height=0.32, color=TEAL, label="今回G1R17 / DLL5実測")
    decorate(ax, [c["case"] for c in old_rows])
    ax.set_xlim(0, 20)
    ax.set_xlabel("主解析時間（分）")
    for i, c in enumerate(old_rows):
        ax.text(c["seconds"] / 60 + 0.15, i - 0.18, f'{c["seconds"] / 60:.2f}', va="center", fontsize=10.5)
        value = current[c["case"]]["analysis_seconds"] / 60
        ax.text(value + 0.15, i + 0.18, f'{value:.2f}', va="center", fontsize=10.5)
    ax.legend(loc="lower left", bbox_to_anchor=(0, 1.015), frameon=False, ncol=2, fontsize=11)
    fig.text(0.07, 0.94, "条件を照合できた10ケース：過去DLL2との参考比較", fontsize=21, weight="bold")
    fig.text(0.07, 0.887, f"時間合計 {old_total / 60:.2f} → {new_total / 60:.2f} 分  ｜  合計比 {old_total / new_total:.3f}倍", fontsize=13)
    fig.text(0.07, 0.838, "物理入力・設定・要素数が一致。保存上下界場のバイナリは10/10で両方とも完全一致。", fontsize=12)
    fig.text(0.07, 0.04, "同一PCの各1回の記録。交互反復測定ではないため、版全体の参考比として扱う。", fontsize=10.5)
    fig.text(0.07, 0.012, "DLL単独の効果・DLLなし比・全26ケースの速度比を示すものではない。", fontsize=10.5)
    save(fig, "prior_dll2_10")


def write_report(summary, previous):
    cases = summary["cases"]
    zip_url = f"{BASE_URL}/raw/refs/heads/main/downloads/{ZIP_NAME}"
    old_total = sum(c["seconds"] for c in previous["cases"])
    current = {c["case"]: c for c in cases}
    new_total = sum(current[c["case"]]["analysis_seconds"] for c in previous["cases"])
    reasons = {"GAP_TARGET_NOT_MET": "GAP未達", "LOWER_ROOT_NOT_VALID": "下界根探索未達",
               "UPPER_ROOT_NOT_VALID": "上界根探索未達", "NATIVE_GLOBAL_INCOMPLETE": "最終確認未完了"}
    lines = []
    for c in cases:
        note = "目標達成" if c["target_met"] else " / ".join(reasons.get(x, x) for x in c["target_unmet_reasons"])
        if c["case"] in {"P08A", "P09A", "P10A"}:
            note = "q8最終HSD予算到達・q4旧場保持" + (" / GAP未達" if c["relative_gap"] > .01 else "")
        q = "/".join(str(x) for x in c["witness_q"])
        elems = "/".join(str(int(x)) for x in c["element_counts"])
        lines.append(f'| [{c["case"]}](data/inputs/{c["case"]}.json) | {c["analysis_seconds"]:.2f} | {c["analysis_seconds"] / 60:.2f} | {c["lower_fs"]:.6f} | {c["upper_fs"]:.6f} | {100 * c["relative_gap"]:.3f} | {q} | {elems} | {note} |')
    table = "\n".join(lines)
    input_lines = "\n".join(f'| [{c["case"]}](data/inputs/{c["case"]}.json) | {c["input_title"]} |' for c in cases)
    text = f"""# G1R17・最新DLL5：間隙水圧を含む26ケースの実測結果

2026年10月12日〜13日（日本時間）に、公開G1R17の同じVBA・DLLを使って26ケースを実Excelで順次解析しました。**全26ケースで独立物理監査済みのFs上下界を取得し、目標精度の達成は10ケースでした。** 主解析時間の合計は **7時間46分7秒（27,966.89秒）** です。粗メッシュの固定Fs比較やPythonの実行時間ではなく、ExcelのFs探索・適応解析の実測値です。

[数値CSV](data/RESULTS.csv) / [処理区間別CSV](data/PHASE_TIMINGS.csv) / [機械可読の結果・監査](data/summary.json) / [採用精度の記録](QUALITY_NOTES.md) / **[結果XLSM26冊・共通DLL・入力・原ログ・保存場・監査コードの証拠ZIP（約83 MiB）]({zip_url})**

## 計算時間

![26ケースの計算時間](figures/runtime_26.png)

最長はP06Nの **6,057.42秒（1時間40分57秒）**、P06Aは **6,049.83秒（1時間40分50秒）** でした。この2ケースで総時間の約43.3%を占めます。中央値は **342.13秒（5.70分）**。今回の測定は各ケース1回です。

## Fs上下界と精度目標

![26ケースのFs上下界](figures/fs_intervals_26.png)

![26ケースの上下界GAP](figures/gap_26.png)

線分は計算上の上下界区間で、統計的な信頼区間ではありません。**Nケースは現行Davis等価関連モデルの結果**であり、非関連流れ則そのものの厳密上下界を保証しません。Aケースは関連流れ則です。MC・Davis変換・Fs定義・許容差を変更していません。

「目標達成」は、上下界GAP≦1%、採用場と根探索の有効・完了判定、独立物理監査を満たすケースです。「Fs取得・目標未達」も監査済み上下界を取得していますが、精密な単一Fsとして扱えません。GAPは `100 × (Fs上界 − Fs下界) / ((Fs上界 + Fs下界)/2)`。下表の数値は表示用に丸めており、CSV・JSONに元の精度で保存しています。

| ケース・入力JSON | 主解析秒 | 分 | Fs下界 | Fs上界 | GAP % | 採用q 下/上 | 採用要素数 下/上 | 判定・未達理由 |
|---|---:|---:|---:|---:|---:|---:|---:|---|
{table}

26ケースの最終実行時エラー番号はすべて0でした。以下の未達を、解析完了や実行時エラーの有無と区別しています。

- **P06A・P06N**：採用q8上下界が独立監査を通過。GAPは8.613%・12.038%で目標未達です。最終採用メッシュは両ケースとも1,000要素・542節点でした。
- **P08A・P09A・P10A**：q8の最終上界確認がHSDの300秒予算に達し、`FINAL_NUMERICAL_UNKNOWN;OLD_WITNESSES_KEPT HSD_SOLVE_BUDGET_EXCEEDED` を記録。監査済みq4上下界を保持しました。P08A・P09AはGAPが1%以内でも、最終確認未完了なので目標達成に数えません。
- **D07A**：GAPは0.978%ですが、全体完了判定は未達。採用qは下/上=8/4です。
- **D01A・D02A・D03N**：上下界場は監査を通りましたが、GAPと根探索の有効判定が未達です。D02Aの採用qは4/8です。

HSD予算はHSD内の開始時刻から反復境界で検査します。準備や通常求解を含む固定Fs全体の所要時間が300秒を超えただけでは、予算到達とは判定していません。途中の`INFEASIBLE`・`UNBOUNDED`はFs探索で利用する判定であり、その記録だけを実行時エラーとは解釈しません。[ケース別の最終棄却・HSD予算記録](data/quality_notes.json)も保存しています。

## 過去DLL2との参考比較

![同条件10ケースの過去DLL2と今回DLL5](figures/prior_dll2_10.png)

過去のDLL2実測と物理入力・公開設定・要素数を照合できた10ケースでは、主時間合計が **{old_total:.2f}秒 → {new_total:.2f}秒**、合計比 **{old_total / new_total:.3f}倍** でした。下界・上界の保存場バイナリは10/10ケースで両方ともSHA-256が完全一致しました。[照合記録](data/prior_dll2_comparison.json)

これは同一PCの各1回の測定を使う**版全体の参考比較**です。交互反復測定ではなく、DLL単独の効果や全26ケースの速度比を示しません。DLLなしの26ケース全解析は再計測していないため、その基準比は掲載していません。過去の「P06A約10時間」も条件を揃えた比較値として扱っていません。

## 計測条件と検証範囲

| 項目 | 実際の条件 |
|---|---|
| 基準コード | G1R17 commit [`24bb5c1`]({BASE_URL}/tree/24bb5c1a09cf7c92997c347ad9f0ee6868e28475)、DLL5改訂、ABI `0x00170005` |
| PC・Excel | AMD Ryzen 9 5900X（12コア・24スレッド）、Windows 11 Pro、64bit Excel 16.0 |
| Fs解析設定 | TARGET=1536、ADAPT=1、CYCLES=3、QUADRATURE=8、FS_TOL=1e-4、GAP=0.01 |
| 公開版の上限 | ADAPT_REFINE_CAP=1000、HSD_SOLVE_SECONDS=300 |
| その他 | 自重固定、kh=0、C0_AUTO=0、RAW_LOWER_AUDIT=1、診断ON、試験注入OFF |
| 計測方法 | 1ケースずつ実Excelで順次実行、各ケース1回。26/26で数値DLL呼出しを記録 |
| 主時間 | 診断ログの `run_end.elapsed_s`。Fs探索・適応解析を含む |
| 別計測 | 初期メッシュ、水圧準備、COM側RPX_Run時間。Excel起動終了・保存・独立Python監査は主時間から除外 |
| 結果保全 | 原本から別名XLSMを作成。製品71モジュールは26/26で不変、追加は記録用`BatchEvidence`のみ |

TARGET・CYCLESは要求設定です。最終要素数や適応の成功回数が必ずその値になるという意味ではありません。最終採用要素数とqは上表に記載しました。処理区間別CSVには入れ子の区間があるため、列を足して主解析時間に一致させることはできません。[計測環境](data/environment.json)

**間隙水圧D08A**は地下水位y=3の固定静水圧 `u=9.81*max(3-y,0)`、有効応力連成ON、サクションOFF、γ=20です。露出面が地下水位より上なので追加の外水圧荷重はありません。反復する浸透FEMの検証ではなく、指定した静水圧場の連成解析です。他25ケースは連成OFFです。節点水圧と独立に計算したP1水圧勾配・`b-grad u` を照合しました。

**D04AのE・D09Aのν**は現行の剛塑性モデルの入力ではありません。両ケースも実際に実行していますが、弾性係数・ポアソン比への感度を確認したケースとは数えません。

**独立Python監査**では、保存済み上下界場を使い、材料・領域面積・外周境界・荷重・自重・水圧勾配、下界のMC降伏・平衡・面力、上界の流れ・速度境界・内部面ジャンプ・接着界面・仕事・目的値を再計算しました。下界は元の物理単位で残差≦1e-7を要求し、上下界端点のλ・仕事の方向も確認しました。これは保存場の独立検査であり、26ケース全てをPythonソルバーで再求解したという意味ではありません。[最終照合記録](data/final_verification.json)

今回の掲載ではソルバー・VBA・DLLを変更していません。測定中に保護対象636ファイルが不変だったことを確認し、元ログ・保存場・ハッシュを証拠ZIPへ収録しました。マクロの信頼設定・セキュリティ・レジストリは変更していません。

## 入力・再現・ダウンロード

ケースIDから今回使用した入力JSONを参照できます。タイトルは元入力の記載です。E・νなどのカタログ項目は上記の適用範囲に従って解釈してください。

<details>
<summary>26ケースの入力タイトル</summary>

| ケース | 入力のタイトル |
|---|---|
{input_lines}

</details>

1. [証拠ZIP]({zip_url})をダウンロードし、展開します。`G1R17_native26/workbooks/` に26冊の結果XLSMと共通DLLがあります。ブックの入力・結果を参照するだけなら再解析は不要です。
2. `G1R17_native26/cases/<ID>/` に実Excelの原ログ・状態・保存応力/速度場・独立監査、`inputs/` に入力、`scripts/` に試験コードを収録しています。
3. 再実行はZIP内 `REPRODUCTION_JA.md` の配置・依存関係に従います。基準commitのリポジトリと `versions/G1R13/vendor` も必要で、ZIP単体では自動試験環境は完結しません。元の測定結果を上書きせず別フォルダーを用います。通常の結果ブック利用にPythonやVBAプロジェクトアクセス許可は不要です。

証拠ZIPのSHA-256：

```text
{ZIP_SHA}  {ZIP_NAME}
```

基準XLSM SHA-256：`{summary['metadata']['source_sha256']}`

基準DLL SHA-256：`{summary['metadata']['dll_sha256']}`

公開用の表・図は記録済み値から生成しています。PythonとNumPy・Matplotlib、日本語フォントを用意し、このフォルダーで `python build_report.py --font <日本語フォントファイル>` を実行すると再生成できます。Excel求解を開始する処理はありません。[掲載内容の照合](PUBLICATION_CHECKS.json)

[G1R17の使い方](../../../versions/G1R17/README.md) / [局所修復の報告](../../../versions/G1R17/REPORT.md) / [リポジトリ先頭](../../../README.md)
"""
    (HERE / "README.md").write_text(text, encoding="utf-8")


def validate(summary, previous):
    cases = summary["cases"]
    assert len(cases) == len({c["case"] for c in cases}) == 26
    assert all(c["fs_obtained"] and c["independent_audit"]["pass"] for c in cases)
    assert sum(c["target_met"] for c in cases) == 10
    assert all(c["error_number"] == 0 and c["dll_actually_used"] for c in cases)
    total = sum(c["analysis_seconds"] for c in cases)
    assert math.isclose(total, summary["analysis_seconds_sum"], abs_tol=1e-8)
    assert math.isclose(total, 27966.886158, abs_tol=1e-8)
    assert all(math.isclose(c["relative_gap"], (c["upper_fs"] - c["lower_fs"]) / ((c["upper_fs"] + c["lower_fs"]) / 2), abs_tol=1e-12) for c in cases)
    with (HERE / "data/RESULTS.csv").open(encoding="utf-8-sig", newline="") as f:
        csv_rows = list(csv.DictReader(f))
    indexed = {c["case"]: c for c in cases}
    assert len(csv_rows) == 26
    for row in csv_rows:
        c = indexed[row["case"]]
        assert float(row["analysis_seconds"]) == c["analysis_seconds"]
        assert float(row["lower_fs"]) == c["lower_fs"]
        assert float(row["upper_fs"]) == c["upper_fs"]
        assert (row["target_met"] == "True") == c["target_met"]
        assert sha256(HERE / "data/inputs" / (c["case"] + ".json")) == c["input_sha256"]
    assert previous["compared_cases"] == previous["both_witnesses_bitwise_equal"] == 10
    assert all(c["input_match"] and not c["profile_differences"] and c["element_counts_match"]
               and all(c["witness_binary_matches"].values()) for c in previous["cases"])
    checks = []
    for file in (HERE / "README.md", REPO / "README.md", REPO / "versions/G1R17/README.md", REPO / "versions/G1R17/REPORT.md"):
        # The publication is responsible for its links, not pre-existing historical references.
        for target in re.findall(r'!?\[[^\]]*\]\(([^)]+)\)', file.read_text(encoding="utf-8-sig")):
            if target.startswith(("http:", "https:", "#")):
                continue
            resolved = (file.parent / target.split("#", 1)[0]).resolve()
            if file.parent == HERE or "G1R17_native26_20261013" in target:
                if resolved == HERE / "PUBLICATION_CHECKS.json":
                    continue  # This record is written after the checks finish.
                assert resolved.exists(), (file, target)
                checks.append(str(resolved.relative_to(REPO)).replace("\\", "/"))
    zip_path = REPO / "downloads" / ZIP_NAME
    assert sha256(zip_path) == ZIP_SHA
    result = {
        "pass": True, "scope": "Publication data/CSV/input hashes/relative links/ZIP hash; no numerical solve repeated",
        "recorded_native_excel_cases": 26, "independently_audited_fs": 26, "target_met": 10,
        "main_seconds_sum": total, "prior_dll2_matched_cases": 10,
        "relative_link_targets_checked": len(checks), "evidence_zip_sha256": ZIP_SHA,
        "figures": {p.name: {"bytes": p.stat().st_size, "sha256": sha256(p)} for p in sorted((HERE / "figures").glob("*"))},
        "data_sha256": {str(p.relative_to(HERE)).replace("\\", "/"): sha256(p) for p in sorted((HERE / "data").rglob("*")) if p.is_file()},
    }
    write_json(HERE / "PUBLICATION_CHECKS.json", result)
    print(json.dumps({k: v for k, v in result.items() if k not in {"figures", "data_sha256"}}, ensure_ascii=True))


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--source", type=Path)
    parser.add_argument("--font")
    args = parser.parse_args()
    if args.source:
        import_records(args.source)
    summary = read_json(HERE / "data/summary.json")
    previous = read_json(HERE / "data/prior_dll2_comparison.json")
    configure_plotting(args.font)
    make_figures(summary, previous)
    write_report(summary, previous)
    validate(summary, previous)


if __name__ == "__main__":
    main()
