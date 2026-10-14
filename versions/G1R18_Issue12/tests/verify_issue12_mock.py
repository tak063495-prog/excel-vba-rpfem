"""Issue #12 control-state mock (pure Python, no Excel).

Mirrors RPX_LowerCache Store/Match/Restore, RPX_RefreshLower and the three q8
metadata gates for the G1R18 base and the Issue12 fix. The source checks assert
that the mirrored conditions are present in the VBA text.
"""
import json, pathlib, sys
ROOT = pathlib.Path(__file__).resolve().parents[1]
BASE = ROOT.parent / "G1R18" / "src" / "vba"
FIX = ROOT / "src" / "vba"

def text(p): return p.read_text(encoding="utf-8")

class Engine:
    def __init__(self, fixed): self.fixed=fixed; self.reset_cache(); self.orig=None; self.q=8; self.epoch=1; self.factor=1.0; self.prepared=False; self.compiles=0; self.hits=0
    def reset_cache(self): self.valid=False; self.key=None; self.saved=None
    def mismatch(self,o):
        if not self.fixed or o is None: return ""
        if o["q"]!=self.q: return "quadrature"
        if o["epoch"]!=self.epoch: return "input_epoch"
        return ""
    def compile(self): self.compiles+=1; self.orig=dict(q=self.q,epoch=self.epoch,factor=self.factor,ready=True)
    def match(self,key):
        if not self.valid: return False
        if self.key!=key: self.reset_cache(); return False
        if self.mismatch(self.saved): self.reset_cache(); return False
        return True
    def store(self,key):
        if self.valid and self.key==key:
            if self.fixed:
                if self.orig and not self.mismatch(self.orig): self.saved=dict(self.orig)
                if self.mismatch(self.saved): self.reset_cache()
            return
        self.reset_cache()
        if self.mismatch(self.orig): return
        self.valid=True; self.key=key; self.saved=dict(self.orig)
    def bound(self,key,factor):
        self.factor=factor
        prepared=self.prepared
        if prepared and self.fixed and self.orig and self.mismatch(self.orig): prepared=False
        hit=False
        if not prepared:
            hit=self.match(key)
            if hit: self.hits+=1; self.orig=dict(self.saved); prepared=True
        if prepared: self.orig["factor"]=factor          # RPX_RefreshLower
        else: self.compile()
        self.prepared=True; self.store(key); return hit
    def gates(self):
        o=self.orig; ok=o["ready"] and o["q"]==8 and o["epoch"]==self.epoch
        return dict(original_dual=ok, accuracy_export=ok and o["factor"]==self.factor, fixed_recovery=ok)

def scenario(fixed):
    e=Engine(fixed); key="model"
    e.q=4
    for f in (1.0,1.2,1.1): e.bound(key,f)
    e.prepared=False  # new Evaluate (ReadInputs resets prepared program)
    e.q=8
    hit=e.bound(key,1.15)
    r=dict(final_hit=hit,original_q=e.orig["q"],factor=e.orig["factor"],gates=e.gates())
    e.bound(key,1.16)  # same-key success repeat
    r["after_repeat_saved_q"]=e.saved["q"] if e.saved else None
    r["after_repeat_gates"]=e.gates()
    f=Engine(fixed); f.q=8; f.bound(key,1.0); f.prepared=False; hit8=f.bound(key,1.1)
    r["q8_fresh_gates"]=f.gates(); r["q8_q8_hit"]=hit8
    return r

def main():
    out={}
    old=scenario(False); new=scenario(True)
    # 1. q4 store -> q8 restore -> refresh keeps q=4 (base) / q=8 (fix)
    assert old["final_hit"] and old["original_q"]==4 and old["factor"]==1.15
    assert new["original_q"]==8 and not new["final_hit"]
    # 2. three gates reject in base, accept in fix
    assert not any(old["gates"].values()) and all(new["gates"].values())
    # 3. same-key repeat keeps q4 in base; q8 in fix
    assert old["after_repeat_saved_q"]==4 and new["after_repeat_saved_q"]==8 and all(new["after_repeat_gates"].values())
    # 4. fresh q8 / q8->q8 cache passes in both, and fix keeps the q8->q8 cache hit
    assert all(old["q8_fresh_gates"].values()) and all(new["q8_fresh_gates"].values()) and new["q8_q8_hit"] and old["q8_q8_hit"]
    b=text(BASE/"RPX_LowerCache.bas"); c=text(FIX/"RPX_LowerCache.bas"); fs=text(FIX/"RPX_Fs.bas")
    for needle,src in [("XOriginalCompiled.subdivisions <> 8",text(BASE/"RPX_Conic.bas")),
                       ("XOriginalCompiled.subdivisions <> 8",text(BASE/"RPX_AccuracyNative.bas")),
                       ("XOriginalCompiled.subdivisions <> 8",text(BASE/"RPX_FixedLowerRecovery.bas")),
                       ("XOriginalCompiled = savedOriginal",c),
                       ("original.subdivisions <> RSubdivisions",c),
                       ("reason = \"original_evidence_\"",c),
                       ("LOWER_CACHE_EVIDENCE_STALE",c),
                       ("savedOriginal = XOriginalCompiled",c),
                       ("XOriginalCompiled.subdivisions <> RSubdivisions",fs)]:
        assert needle in src, needle
    assert "EvidenceMismatch" not in b
    out=dict(status="PASS",base=old,fixed=new)
    print(json.dumps(out,indent=2))
    if len(sys.argv)>1: pathlib.Path(sys.argv[1]).write_text(json.dumps(out,indent=2),encoding="utf-8")
main()
