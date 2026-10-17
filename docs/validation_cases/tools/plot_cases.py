#!/usr/bin/env python3
"""26検証ケースの条件図 (SVG/PNG) と概要グリッドを生成する。
入力: docs/benchmarks/G1R17_native26_20261013/data/inputs/*.json (ベンチマークrunnerが使った入力そのもの)
実行: python docs/validation_cases/tools/plot_cases.py  (リポジトリ直下から)
"""
import json, glob, os
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib import font_manager
from matplotlib.patches import Polygon, FancyArrowPatch

ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", "..", ".."))
INP = os.path.join(ROOT, "docs/benchmarks/G1R17_native26_20261013/data/inputs")
OUT = os.path.join(ROOT, "docs/validation_cases/figures")
for name in ["Noto Sans CJK JP", "IPAexGothic", "Yu Gothic", "MS Gothic"]:
    if any(f.name == name for f in font_manager.fontManager.ttflist):
        plt.rcParams["font.family"] = name; break
plt.rcParams["axes.unicode_minus"] = False
COLORS = ["#d9b38c", "#9ec37a", "#c7c7c7", "#e8d48a"]
BC = {"固定XY": ("#c0392b", "固定XY"), "固定X": ("#2e86c1", "固定X(鉛直ローラー)")}


def load(cid):
    return json.load(open(os.path.join(INP, cid + ".json"), encoding="utf-8"))


def draw(ax, d, small=False):
    s = d["source_case"]
    mats = {m["name"]: m for m in s["materials"]}
    names = [m["name"] for m in s["materials"]]
    for r in s["regions"]:
        m = mats[r["name"]]
        col = COLORS[names.index(r["name"]) % len(COLORS)]
        lab = r["name"] if small else (f'{r["name"]}: c={m["c"]:g} kPa, φ={m["phi"]:g}°, ψ={m["psi"]:g}°, '
                                       f'γ={m["gamma"]:g} kN/m³')
        ax.add_patch(Polygon(r["xy"], closed=True, fc=col, ec="k", lw=0.8, label=lab))
    xs = [p[0] for p in s["outline"]]; ys = [p[1] for p in s["outline"]]
    L = max(max(xs) - min(xs), max(ys) - min(ys))
    used = set()
    for b in d["boundary_explicit"]:
        if b["kind"] in BC:
            col, lab = BC[b["kind"]]
            ax.plot([b["a"][0], b["b"][0]], [b["a"][1], b["b"][1]], color=col, lw=3 if not small else 2,
                    label=None if (small or b["kind"] in used) else lab, solid_capstyle="butt")
            used.add(b["kind"])
        q = b["load0"]
        if q != [0.0, 0.0] and q != [0, 0]:
            (x0, y0), (x1, y1) = b["a"], b["b"]
            n = 6
            for i in range(n + 1):
                x = x0 + (x1 - x0) * i / n; y = y0 + (y1 - y0) * i / n
                ax.add_patch(FancyArrowPatch((x, y + 0.12 * L), (x, y), arrowstyle="-|>", mutation_scale=8,
                                             color="#8e44ad", lw=1))
            ax.plot([x0, x1], [y0 + 0.12 * L] * 2, color="#8e44ad", lw=1,
                    label=None if small else f"上載荷重 {abs(q[1]):g} kPa (鉛直下向き)")
    w = d.get("water", {})
    if w.get("coupling") == 1 and w.get("water_table"):
        wt = w["water_table"]
        ax.plot([p[0] for p in wt], [p[1] for p in wt], "b--", lw=1.5,
                label=None if small else f"地下水位 y={wt[0][1]:g} m（静水圧 u=γw·max(y_w−y,0), γw={w['gamma_w']:g}）")
        ax.text(wt[0][0] + 0.02 * L, wt[0][1] + 0.01 * L, "▽", color="b", fontsize=8)
    ax.set_xlim(min(xs) - 0.04 * L, max(xs) + 0.04 * L)
    ax.set_ylim(min(ys) - 0.04 * L, max(ys) + (0.2 if not small else 0.15) * L)
    ax.set_aspect("equal")
    if not small:
        for x, y in s["outline"]:
            ax.plot(x, y, "k.", ms=3)
            ax.annotate(f"({x:g},{y:g})", (x, y), fontsize=6.5, xytext=(3, 3), textcoords="offset points")


def flow(d):
    return "非関連(Davis等価)" if d["settings"]["ANALYSIS_POLICY"] == "DAVIS_EQUIVALENT" else "関連"


def main():
    os.makedirs(OUT, exist_ok=True)
    ids = sorted(os.path.basename(f)[:-5] for f in glob.glob(os.path.join(INP, "*.json")))
    ids = sorted(ids, key=lambda c: (c[0] != "P", c))
    for cid in ids:
        d = load(cid)
        fig, ax = plt.subplots(figsize=(9, 6))
        draw(ax, d)
        ax.set_title(f'{cid}: {d["source_case"]["title"]}（{flow(d)}, kh={d["settings"]["SEISMIC_KH"]:g}）', fontsize=10)
        ax.set_xlabel("x (m)"); ax.set_ylabel("y (m)")
        ax.legend(fontsize=7, loc="upper right", framealpha=0.9)
        ax.grid(alpha=0.3)
        fig.tight_layout()
        for ext in ("svg", "png"):
            fig.savefig(os.path.join(OUT, f"{cid}.{ext}"), dpi=110)
        plt.close(fig)
    fig, axs = plt.subplots(6, 5, figsize=(20, 20))
    for ax, cid in zip(axs.flat, ids):
        d = load(cid); draw(ax, d, small=True)
        ax.set_title(f'{cid}（{flow(d)}）', fontsize=10); ax.tick_params(labelsize=6)
    for ax in list(axs.flat)[len(ids):]:
        ax.axis("off")
    fig.suptitle("検証26ケースの形状と材料区分（赤=固定XY, 青=固定X, 青破線=地下水位, 紫=上載荷重）", fontsize=14)
    fig.tight_layout(rect=(0, 0, 1, 0.97))
    for ext in ("svg", "png"):
        fig.savefig(os.path.join(OUT, f"overview.{ext}"), dpi=80)


if __name__ == "__main__":
    main()
