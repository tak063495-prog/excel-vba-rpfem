"""Map Python sim steps to estimated VBA (Winmax2) wall time.
Calibration (Winmax2, Fix8 rc2 logs):
  DLL upper : 3.45 s @ 22741 cones (P02A_THIN), ~linear^1.2
  DLL lower : 2.55 s @ 6198 cones (P02A_THIN)
  VBA overhead per solve: upper 6.9 s @ 22741 cones, lower 13.3 s @ 6198 cones (linear)
  HSD upper : 20 s + 3.7 ms/cone (P10A rc2: 41k->180 s, 66k->261 s, 90k->353 s)
  HSD lower : assumed 4x DLL lower (uncalibrated)
  VBA solves per root step: 13 (P02A/P05A logs: 12-13 per new mesh)
Scenarios:
  S0 as-is: DLL if est<=5 s (DIRECT_DLL_SECONDS) else HSD; HSD > 300 s -> step fails (chain stops)
  S1: DIRECT_DLL_SECONDS scaled (no 5 s cutoff); 25% of upper solves fail and fall to HSD
  S2: S1 with near-root DLL failures fixed (0%)
"""
import json,sys,glob,os
def dll(side,c): return 3.45*(c/22741)**1.2 if side=='upper' else 2.55*(c/6198)**1.2
def ovh(side,c): return 6.9*c/22741 if side=='upper' else 13.3*c/6198
def hsd(side,c): return 20+0.0037*c if side=='upper' else 4*dll(side,c)
def solve(side,c,sc):
    d=dll(side,c)
    if sc=='S0':
        if d<=5: return ovh(side,c)+d+(0.25*hsd(side,c) if side=='upper' else 0), True
        h=hsd(side,c); return ovh(side,c)+min(h,300), h<=300
    f=0.25 if (sc=='S1' and side=='upper') else 0.0
    return ovh(side,c)+d+f*min(hsd(side,c),450), True
NS=13
for f in sorted(glob.glob(sys.argv[1]+'/*/log.jsonl')):
    name=os.path.basename(os.path.dirname(f))
    steps=[json.loads(l) for l in open(f)]
    steps=[e for e in steps if e['event']=='step' and e.get('cones')]
    out=[name]
    for sc in ('S0','S1','S2'):
        T=0;gap=None;stopped=None;bestL=bestU=None
        for e in steps:
            t,ok=solve(e['side'],e['cones'],sc)
            T+=NS*t
            if not ok: stopped=f"{e['side'][0]}{e['round']}@{e['elements']}"; break
            gap=e['gap']
        out.append(f"{sc}: {T/3600:.2f}h gap={(gap or 0)*100:.2f}%"+(f" STOP {stopped}" if stopped else ""))
    print(' | '.join(out))
