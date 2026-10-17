"""入力外形から斜面の寸法を算出する（算出値）。"""
import math


def geometry(outline):
    xs = [p[0] for p in outline]; ys = [p[1] for p in outline]
    xmin, xmax, ymin, ymax = min(xs), max(xs), min(ys), max(ys)
    i0 = next(i for i, p in enumerate(outline) if p[0] == xmax and p[1] == ymin)
    surf = []
    for p in outline[i0 + 1:] + outline[:0]:
        surf.append(tuple(p))
        if p[0] == xmin:
            break
    if surf and surf[0][0] == xmax and len(surf) > 1 and surf[0][1] > ymin:
        pass
    if not surf or surf[0][0] != xmax:
        surf.insert(0, (xmax, ymin))
    # 右から: 法尻側平坦部
    k = 0
    while k + 1 < len(surf) and surf[k + 1][1] == surf[k][1]:
        k += 1
    toe = surf[k]
    j = len(surf) - 1
    while j - 1 >= 0 and surf[j - 1][1] == surf[j][1]:
        j -= 1
    crest = surf[j]
    face = surf[k:j + 1]
    segs = []
    for a, b in zip(face, face[1:]):
        dx, dy = a[0] - b[0], b[1] - a[1]
        if dy != 0:
            segs.append(dict(a=a, b=b, run=dx, rise=dy, beta=math.degrees(math.atan2(dy, dx)), n=dx / dy))
    H = crest[1] - toe[1]; run = toe[0] - crest[0]
    return dict(xmin=xmin, xmax=xmax, ymin=ymin, ymax=ymax, W=xmax - xmin, Ht=ymax - ymin, toe=toe, crest=crest,
                H=H, run=run, beta=math.degrees(math.atan2(H, run)), n=run / H if H else float("nan"),
                crest_w=crest[0] - xmin, toe_w=xmax - toe[0], fdepth=toe[1] - ymin, segs=segs)
