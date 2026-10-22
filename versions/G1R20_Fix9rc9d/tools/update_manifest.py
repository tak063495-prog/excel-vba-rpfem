"""Refresh SOURCE_MANIFEST.json for Fix9 rc6 from src/vba (import = CP932 encoding of the UTF-8 source)."""
from pathlib import Path
import hashlib,json,shutil
ROOT=Path(__file__).resolve().parents[1];BASE=ROOT.parent/'G1R20_Fix9rc2'
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
base=json.loads((BASE/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
m=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
changed=[]
for name in sorted(set(base['modules'])|{p.name for p in (ROOT/'src/vba').glob('*.bas')}):
    src=ROOT/'src/vba'/name
    if name in base['modules'] and sha(src)==base['modules'][name]['utf8_sha256']:
        m['modules'][name]=dict(base['modules'][name]);assert sha(ROOT/'import'/name)==base['modules'][name]['cp932_sha256'];continue
    (ROOT/'import'/name).write_bytes(src.read_bytes().decode('utf8').encode('cp932'))
    m['modules'][name]={'utf8_sha256':sha(src),'cp932_sha256':sha(ROOT/'import'/name)};changed.append(name)
m.update(schema='rpfem.g1r20.fix9rc7b.sources.v1',version='G1R20_Fix9rc9d_candidate',base_version='G1R20_Fix9rc2',
    base_manifest_sha256=sha(BASE/'SOURCE_MANIFEST.json'),changed_modules=changed,
    base_module_sha256={k:v['utf8_sha256'] for k,v in base['modules'].items()},
    scope='G1R20 Fix9 rc9d candidate = rc9c + ACCURACY_DLL_UPPER_EQUILIBRATE default 0 (code default, settings row, both presets; P02A_THIN eq0 PASS 0.457% vs eq1 0.531%). rc9c = rc9b + midnight-safe monotonic elapsed clock (new RPX_Clock; all elapsed Timer uses; default-off TEST_CLOCK wrap hook). rc9b = rc9a + upper original primal DLL equilibrate (ACCURACY_DLL_UPPER_EQUILIBRATE default 1, 0 = rc9a) + phase-specific dual/call-id diagnostic logging (no stale dual fields); no acceptance/audit/tolerance change. rc9a = rc8 + bug fix: thin-budget failure with no iterate treated as numerical UNKNOWN (no recovery on missing iterate, guarded array probes) and confirmed best lower/upper preserved and written on late fatal errors (adapt_fatal_preserved), default-off test hooks. rc8 = rc7 + thin-only same-mesh HSD warm start (raw x/y/s/z/tau/kappa cache keyed by exact layout/model, recentred blend .2, iteration-0 reject x10, probe 15 its, one cold retry within same budget/wall; cross-mesh not implemented). rc7 = rc6c + thin-only: large-upper HSD floor 1800 s/600 it with hard wall incl. grace, lower HSD floor 1200 s (>=1400 el or extra), narrow directional Fs seed bracket from lineage hint (~9 evals), cost-ranked upper marks when cap binds, 1 targeted extra lower, protection reserves planned extra lower. rc6c = rc6 with ADAPT_THIN_NOMINAL_SECONDS default 18171 (user decision). rc6 (Codex-written, reviewed) = rc5 (+#38) + thin-chain-only: upper round element cap 3000/growth 1.6 with fitted marks, large thin upper DLL status-5 retry (700 it, equilibration) and larger HSD budget/iterations, real case clock via ADAPT_THIN_CASE_LIMIT_SECONDS, up to 3 time-funded extra lower rounds with upper-protect. rc5 = rc4b + thin-chain-only robustness: DLL status 5 as upper below-root sign (prior dual guide gate), keep fully valid best of failed round, thin HSD floor from remaining time (optional deadline or ADAPT_THIN_NOMINAL_SECONDS 14400), lower large-mesh theta factor 0.5, one extra retry per side. rc4b: rc4 with ADAPT_THIN_EARLY_UPPER_THETA default 0.95 on top of rc3b (Codex-written, reviewed): after an early start is taken, upper chain uses ADAPT_THIN_EARLY_UPPER_THETA=0.95 and ADAPT_THIN_EARLY_UPPER_CAP=7000 (budget check uses them too); otherwise identical to rc3b. rc3b: conditional early start ADAPT_THIN_EARLY_START=1 (Fix9ThinEarlyBudget: min upper rounds 10, min lower rounds 3, lower end <= upper rounds, theta >= 0.8/0.9, cap/coarse >= 2, remaining optional deadline >= 3600 s if active; 2 = always as rc2, 0 = off); ADAPT_THIN_AUX_ALMOST=1; lower cap 6000',
    validation_warning='Candidate; build/compile only; no VBA case run yet; not published')
payload={k:m[k] for k in ('modules','workbook_modules')}
m['source_bundle_sha256']=hashlib.sha256(json.dumps(payload,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()
(ROOT/'SOURCE_MANIFEST.json').write_text(json.dumps(m,indent=2,ensure_ascii=False)+'\n',encoding='utf8')
print('changed',changed)
