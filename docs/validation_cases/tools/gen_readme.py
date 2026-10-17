#!/usr/bin/env python3
"""docs/validation_cases/README.md を入力JSONとTest_Case_Catalog.mdから生成する。"""
import json, os, re
ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
INP = os.path.join(ROOT, "docs/benchmarks/G1R17_native26_20261013/data/inputs")
CAT = os.path.join(ROOT, "validation/G1R12_native25_20261009/Test_Case_Catalog.md")
cat = open(CAT, encoding="utf-8").read()
secs = {m.group(1): m.group(0) for m in re.finditer(r"### (\w{4}) — .*?(?=\n### |\n## |\Z)", cat, re.S)}
lst = {r[0]: r for r in re.findall(r"^\| (\w{4}) \| (.*?) \| (\w+) \| (.*?) \|$", cat, re.M)}

TEST = {  # 試験目的: カタログの「使用順序」「由来区分」「タイトル」から要約（独自の推測は入れない）
 "P": "文献ケース（R02）の再現。Nは文献値との比較、Aは関連流れ則の双子モデル",
}
PURPOSE = {
 "P01":"均質斜面 Case1（低φ=5°）。文献Fs≈0.65〜0.70","P02":"均質斜面 Case2（φ=25°）。文献Fs≈1.30〜1.36",
 "P03":"均質斜面 Case3。文献Fs≈1.06〜1.20","P04":"均質斜面 Case4（高Fs）。文献Fs≈2.24〜2.34",
 "P05":"3層水平成層斜面（保留評価群）","P06":"薄い摩擦性弱層（c=0, φ=25°）を含む斜面（保留評価群）",
 "P07":"Griffiths–Lane Example1（D=1、基礎なし）","P08":"均質非排水粘土（φ=0, Tresca相当）",
 "P09":"軟弱粘土基礎（基礎cu=30 < 斜面cu=50、保留評価群）","P10":"硬い粘土基礎（基礎cu=75）で破壊機構が変わる場合（保留評価群）",
 "D01":"c=0・45°法面・φ=30°（上界の適格性検査群）","D02":"c=0・勾配2:3・φ=45°（表層滑り極限機構 tan45°/tan(arctan(2/3))=1.5 との照合）",
 "D03":"D02と同形状の非関連（ψ=0）版（異常系/不変性）","D04":"弾性係数E/10の感度（Eは剛塑性モデルの入力ではないが実際に実行する）",
 "D05":"幾何相似（長さ×1000, γ/1000）。理想FsはP02Aと同じ","D06":"平行移動（+1000,+2000）。FsはP02Aと同じはず",
 "D07":"天端上載荷重20 kPa（x=2〜8）","D08":"固定静水圧間隙水圧場（水位y=3）",
 "D09":"ほぼ非圧縮 ν=0.49（νは剛塑性モデルの入力ではないが実際に実行する）",
}
def f(x): return f"{x:g}"
ids = sorted([n[:-5] for n in os.listdir(INP) if n.endswith(".json")], key=lambda c: (c[0] != "P", c))
rows, body = [], []
for cid in ids:
    d = json.load(open(os.path.join(INP, cid + ".json"), encoding="utf-8")); s = d["source_case"]
    xs = [p[0] for p in s["outline"]]; ys = [p[1] for p in s["outline"]]
    W, H = max(xs) - min(xs), max(ys) - min(ys)
    mats = "; ".join(f'{m["name"]} c={f(m["c"])} φ={f(m["phi"])} ψ={f(m["psi"])} γ={f(m["gamma"])}' for m in s["materials"])
    fl = "非関連（Davis等価）" if d["settings"]["ANALYSIS_POLICY"] == "DAVIS_EQUIVALENT" else "関連（ψ=φ）"
    w = d["water"]
    water = f'静水圧 水位y={f(w["water_table"][0][1])} m' if w.get("coupling") == 1 else "なし"
    sec = secs.get(cid, "")
    refs = re.findall(r"^\| (?!材料|報告|---)([^|]+?) \| ([0-9.]+) \| ([^|]+?) \|$", sec, re.M)
    reff = " / ".join(sorted({r[1] for r in refs}, key=float)) if refs else "記載なし"
    if refs: reff = f"{min(float(r[1]) for r in refs):g}〜{max(float(r[1]) for r in refs):g}" if len(refs) > 1 else refs[0][1]
    load = "天端 20 kPa（x=2〜8）" if s.get("surcharge") else "なし"
    rows.append(f"| [{cid}](#{cid.lower()}) | {lst[cid][1]} | 幅{f(W)}×高{f(H)} m | {mats} | {water} | {load} | {fl} | {reff} | {PURPOSE[cid[:3]]} |")
    src = re.search(r"- 出典箇所: (.*)", sec); kind = re.search(r"- 由来区分: `(.*?)`", sec)
    notes = re.search(r"補足仮定・制限:\n(.*?)(?=\n\n|\n[^-]|\Z)", sec, re.S)
    extra = [l for l in sec.splitlines() if l.startswith(("不変性試験", "tan45", "追加荷重", "間隙水圧"))]
    b = [f"### {cid}", "", f"![{cid}](figures/{cid}.svg)", "",
         f"- 名称: {lst[cid][1]}", f"- 目的: {PURPOSE[cid[:3]]}",
         f"- 文献/由来: {lst[cid][3]}（出典箇所: {src.group(1) if src else '記載なし'}）",
         f"- 外形（反時計回り, m）: `{' → '.join(f'({f(x)}, {f(y)})' for x, y in s['outline'])}`",
         f"- 面積: {f(s['area'])} m² ／ 総重量: {f(d['total_weight'])} kN/m ／ メッシュ目標h (M0/M1/M2): {' / '.join(f(h) for h in s['h'])} m",
         f"- 流れ則: {fl}（ψ政策 `{s['policy']}`）", f"- 水: {water} ／ 地震: kh={f(d['settings']['SEISMIC_KH'])}（なし） ／ 荷重: {load}",
         "- 境界: 底面 固定XY、左右鉛直端 固定X（鉛直ローラー）、その他 自由" + ("（右鉛直端なし）" if cid.startswith("P07") else ""), "",
         "| 材料 | E (kPa) | ν | γ (kN/m³) | c (kPa) | φ (°) | ψ (°) |", "|---|---:|---:|---:|---:|---:|---:|"]
    b += [f'| {m["name"]} | {f(m["E"])} | {f(m["nu"])} | {f(m["gamma"])} | {f(m["c"])} | {f(m["phi"])} | {f(m["psi"])} |' for m in s["materials"]]
    b.append("")
    if refs:
        b += ["参照Fs（文献の方法別報告値。本メッシュの保証値ではない）:", "", "| 方法 | Fs | 出典 |", "|---|---:|---|"]
        b += [f"| {r[0]} | {r[1]} | {r[2].replace('; secondary when other authors are named', '')} |" for r in refs] + [""]
    else:
        b += ["参照Fs: 記載なし（派生/診断入力のため元ケースの値をコピーしていない）", ""]
    for e in extra: b.append(f"- {e}")
    if notes and notes.group(1).strip():
        b += ["補足仮定・制限（カタログ原文）:", "", notes.group(1).strip()]
    body += b + [""]

head = """# 検証26ケースの条件（図説）

このページは、G1R12 以降の速度・精度ベンチマーク（例: [G1R17 native26](../benchmarks/G1R17_native26_20261013/README.md)）で使っている26の検証ケースの条件を、図と表でまとめたものです。図と表は、runner が実際に読む入力JSON（[`docs/benchmarks/G1R17_native26_20261013/data/inputs/`](../benchmarks/G1R17_native26_20261013/data/inputs/)）と、ケース定義書 [`Test_Case_Catalog.md`](../../validation/G1R12_native25_20261009/Test_Case_Catalog.md) から自動生成しています（[`tools/`](tools/)）。

## 共通条件

- 単位 m, kN, kPa, 度。平面ひずみ、単位奥行 1 m。Mohr–Coulomb 材料。強度低減法で安全率 Fs の下界・上界を求める。
- 境界: 底面は固定XY（赤）、左右の鉛直端は固定X＝水平変位拘束（青）、その他の露出面は自由。P07 は右鉛直端を持たない。
- 地震（水平震度 kh）は全ケース 0。間隙水圧は D08A のみ。上載荷重は D07A のみ。追加の引張カットオフはなし。
- 図の材料色: ケース内の材料順に 茶→緑→灰。P08A は斜面と基礎が同じ強度だが領域を分けて表示している。

## 命名規則

- 先頭文字: **P** = 文献（Paper）由来のケース、**D** = 独自設計の診断（Diagnostic）ケース。番号は形状・材料の組。
- 末尾文字: **A** = 関連流れ則（ψ=φ）。**N** = 非関連流れ則（文献のψ。ソルバー内では Davis 等価の関連モデル `DAVIS_EQUIVALENT` で扱う）。
- 文献記号: R02 = Jia ほか 2024（eSPFEM、Cheng 2007 などの比較表を含む）、R03 = Griffiths & Lane（Example1〜4）、R10・R11 = P06 の形状照合に使った文献。正式な書誌情報は repo 内に見当たらないため記載していない（カタログでは `Source_Audit.md / sources.json` を参照としているが未収録）。
- 由来区分: `literature_geometry_parameters`=文献の形状・定数そのまま、`derived_associated_twin`=N ケースの関連版（参照Fsはコピーしない）、`nondimensional_literature_case_dimensionalized`=無次元の文献ケースに寸法を与えたもの、`designed_diagnostic`=診断用の設計ケース。

## 全体図

![全26ケース](figures/overview.svg)

## 一覧

| ID | 名称 | 形状 | 材料（c kPa, φ°, ψ°, γ kN/m³） | 水 | 荷重 | 流れ則 | 参照Fs | 目的 |
|---|---|---|---|---|---|---|---|---|
"""
tail = """
## 不明・記載なし

- R02/R03/R10/R11 の正式な書誌情報（著者・誌名・年・DOI）。
- A ケースと D ケースの参照Fs（D02A の 1.5 は極限機構の照合値で、有限メッシュの保証値ではない）。P03N の文献 Case3 の土質定数の出典表記以外の詳細。
- 各ケースの「目標精度」はケース定義には含まれない（ベンチマークの設定 GAP による）。
- 入力JSON内の `settings`（CYCLES など）は G1R17 の計測プロファイルの値で、各版の製品設定とは異なる場合がある。

再生成: リポジトリ直下で `python docs/validation_cases/tools/plot_cases.py && python docs/validation_cases/tools/gen_readme.py`（matplotlib と日本語フォント Noto Sans CJK JP 等が必要）。
"""
open(os.path.join(ROOT, "docs/validation_cases/README.md"), "w", encoding="utf-8").write(head + "\n".join(rows) + "\n\n## ケース別の条件\n\n" + "\n".join(body) + tail)
