"""Compare G1R19_Issue11 native case results with G1R19 (same runner/settings) and G1R17 DLL5.
usage: python compare_cases.py NEW_DIR G1R19_DIR G1R17_SUMMARY_JSON OUT_JSON"""
import json,sys,csv
from pathlib import Path
new,old,g17,outp=Path(sys.argv[1]),Path(sys.argv[2]),Path(sys.argv[3]),Path(sys.argv[4])
ORDER=['D08A','P04A','P04N','P02N','D07A','P08A','P09A','P07A','P07N','P05A']
g=json.loads(g17.read_text(encoding='utf8'));g={c['case']:c for c in g['cases']}
def pick(r):
    if not r:return None
    a=r.get('independent_audit') or {}
    ls=r.get('log_summary') or {}
    return {'status':r.get('status'),'lower_fs':r.get('lower_fs'),'upper_fs':r.get('upper_fs'),'gap_pct':None if r.get('relative_gap') is None else 100*r['relative_gap'],
            'analysis_seconds':r.get('analysis_seconds'),'target_met':r.get('target_met'),'audit_pass':a.get('pass') if isinstance(a,dict) else None,
            'reasons':r.get('target_unmet_reasons'),'solve_count':ls.get('solve_count') if isinstance(ls,dict) else None}
def events(r):
    p=r and r.get('diagnostic_path')
    if not p or not Path(p).exists():return None
    rows=list(csv.DictReader(open(p,encoding='utf-16'),delimiter='\t'))
    return {'final_evaluations':sum(1 for x in rows if x['event']=='adapt_eval' and x['detail'].startswith('label=final;')),
            'terminal_pair_reuse':sum(1 for x in rows if x['event']=='terminal_pair_reuse'),
            'terminal_pair_reuse_rejected':sum(1 for x in rows if x['event']=='terminal_pair_reuse_rejected')}
res=[]
for c in ORDER:
    load=lambda d:json.loads((d/'results'/f'{c}.json').read_text(encoding='utf8')) if (d/'results'/f'{c}.json').exists() else None
    n,o=load(new),load(old)
    res.append({'case':c,'issue11':pick(n),'issue11_events':events(n),'g1r19':pick(o),'g1r19_events':events(o),'g1r17':pick(g.get(c))})
outp.write_text(json.dumps(res,indent=1,ensure_ascii=False),encoding='utf8')
f=lambda v,k:'-' if not v or v.get(k) is None else (f"{v[k]:.4f}" if isinstance(v[k],float) and k!='analysis_seconds' else (f"{v[k]:.0f}" if isinstance(v[k],float) else str(v[k])))
for r in res:
    print(r['case'],*[f"{k}: {f(r[k],'lower_fs')}-{f(r[k],'upper_fs')} gap {f(r[k],'gap_pct')}% {f(r[k],'analysis_seconds')}s tgt={f(r[k],'target_met')} audit={f(r[k],'audit_pass')}" for k in ('issue11','g1r19','g1r17')],r['issue11_events'],r['g1r19_events'],sep=' | ')
