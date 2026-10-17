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
from matplotlib.patches import Polygon, FancyArrowPatch, Arc
from matplotlib.path import Path as MplPath
import sys
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from geom import geometry

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


def _inside_point(xy):
    path = MplPath(xy)
    cx = sum(p[0] for p in xy) / len(xy); cy = sum(p[1] for p in xy) / len(xy)
    if path.contains_point((cx, cy)):
        return cx, cy
    xs = [p[0] for p in xy]; ys = [p[1] for p in xy]
    best = None
    for i in range(41):
        for j in range(41):
            x = min(xs) + (max(xs) - min(xs)) * i / 40; y = min(ys) + (max(ys) - min(ys)) * j / 40
            if path.contains_point((x, y)):
                dd = (x - cx) ** 2 + (y - cy) ** 2
                if best is None or dd < best[0]:
                    best = (dd, x, y)
    return best[1], best[2]


def _area(xy):
    return abs(sum(xy[i][0] * xy[i - 1][1] - xy[i - 1][0] * xy[i][1] for i in range(len(xy)))) / 2


def _dim(ax, p0, p1, text, off, vertical, fs):
    (x0, y0), (x1, y1) = p0, p1
    if vertical:
        ax.annotate("", (off, y0), (off, y1), arrowprops=dict(arrowstyle="<->", lw=0.8, color="k"))
        for y, x in ((y0, x0), (y1, x1)):
            ax.plot([x, off], [y, y], color="gray", lw=0.5, ls=":")
        ax.text(off, (y0 + y1) / 2, " " + text, fontsize=fs, va="center", ha="left", rotation=90,
                bbox=dict(fc="white", ec="none", pad=0.5, alpha=0.85))
    else:
        ax.annotate("", (x0, off), (x1, off), arrowprops=dict(arrowstyle="<->", lw=0.8, color="k"))
        for x, y in ((x0, y0), (x1, y1)):
            ax.plot([x, x], [y, off], color="gray", lw=0.5, ls=":")
        ax.text((x0 + x1) / 2, off, text, fontsize=fs, ha="center", va="center",
                bbox=dict(fc="white", ec="none", pad=0.5, alpha=0.85))


def draw(ax, d, small=False):
    s = d["source_case"]
    mats = {m["name"]: m for m in s["materials"]}
    names = [m["name"] for m in s["materials"]]
    nonassoc = d["settings"]["ANALYSIS_POLICY"] == "DAVIS_EQUIVALENT"
    g = geometry(s["outline"])
    L = max(g["W"], g["Ht"]); x0, y0 = g["xmin"], g["ymin"]
    def u(v):  # 寸法の表示
        return f"{v:g} m"
    tot = sum(_area(r["xy"]) for r in s["regions"])
    for r in s["regions"]:
        m = mats[r["name"]]
        col = COLORS[names.index(r["name"]) % len(COLORS)]
        ax.add_patch(Polygon(r["xy"], closed=True, fc=col, ec="k", lw=0.8))
        lab = f'{r["name"]}\nc={m["c"]:g} kPa, φ={m["phi"]:g}°\nγ={m["gamma"]:g} kN/m³' + (f', ψ={m["psi"]:g}°' if nonassoc else "")
        if small:
            lab = f'c={m["c"]:g} φ={m["phi"]:g}'
        px, py = _inside_point(r["xy"])
        thin = _area(r["xy"]) < 0.08 * tot
        if thin and not small:
            ax.annotate(lab, (px, py), xytext=(px + 0.12 * L, py - 0.18 * L), fontsize=7.5, ha="center",
                        arrowprops=dict(arrowstyle="->", lw=0.8), bbox=dict(fc="white", ec=col, lw=1.5, pad=2))
        else:
            ax.text(px, py, lab, fontsize=7.5 if not small else 6.5, ha="center", va="center",
                    bbox=dict(fc="white", ec="none", alpha=0.75, pad=1.5))
    used = set()
    for b in d["boundary_explicit"]:
        if b["kind"] in BC:
            col, lab = BC[b["kind"]]
            ax.plot([b["a"][0], b["b"][0]], [b["a"][1], b["b"][1]], color=col, lw=3 if not small else 2,
                    label=None if (small or b["kind"] in used) else lab, solid_capstyle="butt")
            used.add(b["kind"])
        q = b["load0"]
        if q != [0.0, 0.0] and q != [0, 0]:
            (qa, qy), (qb, qy2) = b["a"], b["b"]
            for i in range(7):
                x = qa + (qb - qa) * i / 6
                ax.add_patch(FancyArrowPatch((x, qy + 0.1 * L), (x, qy), arrowstyle="-|>", mutation_scale=8,
                                             color="#8e44ad", lw=1))
            ax.plot([qa, qb], [qy + 0.1 * L] * 2, color="#8e44ad", lw=1, label=None if small else "上載荷重")
            ax.text((qa + qb) / 2, qy + 0.115 * L, f"q={abs(q[1]):g} kPa" + ("" if small else f"（x={min(qa,qb):g}〜{max(qa,qb):g} m, 幅{abs(qb-qa):g} m）"),
                    color="#8e44ad", fontsize=7.5 if not small else 6, ha="center", va="bottom")
    w = d.get("water", {})
    if w.get("coupling") == 1 and w.get("water_table"):
        wt = w["water_table"]
        ax.plot([p[0] for p in wt], [p[1] for p in wt], "b--", lw=1.5, label=None if small else "地下水位")
        ax.text(wt[0][0] + 0.02 * L, wt[0][1] + 0.01 * L, f"▽ GWL y={wt[0][1]:g} m" + ("" if small else f"（静水圧, γw={w['gamma_w']:g} kN/m³）"),
                color="b", fontsize=7.5 if not small else 6, va="bottom")
    # 法面勾配
    fs = 7.5 if not small else 6
    t, c = g["toe"], g["crest"]
    seg0 = g["segs"][0]
    rad = 0.07 * L
    ax.add_patch(Arc(t, 2 * rad, 2 * rad, theta1=180 - seg0["beta"], theta2=180, color="#c0392b", lw=1))
    ax.plot([t[0], t[0] + 0.12 * L], [t[1], t[1]], color="gray", lw=0.5, ls=":") if g["toe_w"] == 0 else None
    bt = f'β={g["beta"]:.1f}°（1:{g["n"]:g}）'
    if len(g["segs"]) > 1 and not small:
        bt = f'全体 β={g["beta"]:.1f}°（1:{g["n"]:g}）'
        for sg in g["segs"]:
            mx = (sg["a"][0] + sg["b"][0]) / 2; my = (sg["a"][1] + sg["b"][1]) / 2
            ax.text(mx + 0.02 * L, my + 0.02 * L, f'{sg["beta"]:.1f}°（1:{sg["n"]:.3g}）', fontsize=fs, color="#c0392b",
                    bbox=dict(fc="white", ec="none", alpha=0.8, pad=0.5))
    if not small: ax.text(t[0] - rad * 0.9, t[1] - 0.012 * L, bt, fontsize=fs, color="#c0392b", ha="right", va="top",
            bbox=dict(fc="white", ec="none", alpha=0.8, pad=0.5))
    # 寸法線
    top = g["ymax"] + (0.06 if small else (0.2 if s.get("surcharge") else 0.05)) * L
    if not small:
        if g["crest_w"] > 0: _dim(ax, (x0, g["ymax"]), c, f'天端 {u(g["crest_w"])}', top, False, fs)
        _dim(ax, c, (t[0], t[1]), f'法面水平長 {u(g["run"])}', top, False, fs)
        if g["toe_w"] > 0: _dim(ax, t, (g["xmax"], t[1]), f'法尻側 {u(g["toe_w"])}', top, False, fs)
        _dim(ax, (x0, y0), (g["xmax"], y0), f'全幅 W={u(g["W"])}', y0 - 0.07 * L, False, fs)
        _dim(ax, (x0, y0), (x0, g["ymax"]), f'全高 {u(g["Ht"])}', x0 - 0.07 * L, True, fs)
    xr = g["xmax"] + 0.05 * L
    _dim(ax, t, c, f'H={u(g["H"])}', xr, True, fs)
    if g["fdepth"] > 0 and not small:
        _dim(ax, (g["xmax"], y0), (g["xmax"], t[1]), f'基礎 {u(g["fdepth"])}', xr, True, fs)
    ax.set_xlim(x0 - (0.11 if not small else 0.04) * L, g["xmax"] + (0.12 if not small else 0.12) * L)
    ax.set_ylim(y0 - (0.11 if not small else 0.04) * L, g["ymax"] + (0.25 if not small else 0.14) * L)
    ax.set_aspect("equal")
    if not small:
        for x, y in s["outline"]:
            ax.plot(x, y, "k.", ms=3)
            ax.annotate(f"({x:g},{y:g})", (x, y), fontsize=6, xytext=(3, -9), textcoords="offset points", color="#444")


def flow(d):
    return "非関連(Davis等価)" if d["settings"]["ANALYSIS_POLICY"] == "DAVIS_EQUIVALENT" else "関連"


def main():
    os.makedirs(OUT, exist_ok=True)
    ids = sorted(os.path.basename(f)[:-5] for f in glob.glob(os.path.join(INP, "*.json")))
    ids = sorted(ids, key=lambda c: (c[0] != "P", c))
    for cid in ids:
        d = load(cid)
        fig, ax = plt.subplots(figsize=(10, 7))
        draw(ax, d)
        ax.set_title(f'{cid}: {d["source_case"]["title"]}（{flow(d)}, kh={d["settings"]["SEISMIC_KH"]:g}）', fontsize=10)
        ax.set_xlabel("x (m)"); ax.set_ylabel("y (m)")
        ax.legend(fontsize=7, loc="upper center", bbox_to_anchor=(0.5, -0.1), ncol=4, framealpha=0.9)
        ax.grid(alpha=0.3)
        fig.tight_layout()
        for ext in ("svg", "png"):
            fig.savefig(os.path.join(OUT, f"{cid}.{ext}"), dpi=110)
        plt.close(fig)
    fig, axs = plt.subplots(6, 5, figsize=(20, 20))
    for ax, cid in zip(axs.flat, ids):
        d = load(cid); draw(ax, d, small=True)
        gg = geometry(d["source_case"]["outline"]); ax.set_title(f'{cid}（{flow(d)}） H={gg["H"]:g} m, β={gg["beta"]:.1f}°', fontsize=10); ax.tick_params(labelsize=6)
    for ax in list(axs.flat)[len(ids):]:
        ax.axis("off")
    fig.suptitle("検証26ケースの形状と材料区分（赤=固定XY, 青=固定X, 青破線=地下水位, 紫=上載荷重, c kPa・φ°）", fontsize=14)
    fig.tight_layout(rect=(0, 0, 1, 0.97))
    for ext in ("svg", "png"):
        fig.savefig(os.path.join(OUT, f"overview.{ext}"), dpi=110)


if __name__ == "__main__":
    main()
