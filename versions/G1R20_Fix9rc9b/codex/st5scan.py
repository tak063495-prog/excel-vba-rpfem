import sys,re,io
def rd(p):
    b=open(p,'rb').read()
    return b.decode('utf-16') if b[:2] in (b'\xff\xfe',b'\xfe\xff') else b.decode('utf-8','replace')
def kv(s): return dict(x.split('=',1) for x in s.split(';') if '=' in x)
for p in sys.argv[1:]:
    rows=[l.split('\t') for l in rd(p).splitlines()[1:]]
    rows=[r for r in rows if len(r)>=11]
    cur={};out=[];el={}
    for r in rows:
        sid=r[2];ev=r[6];d=r[10]
        if ev=='solve_start': el[sid]=kv(d).get('elements')
        if ev=='direct_dll_end':
            k=kv(d)
            if k.get('mode')=='upper': cur[sid]=dict(fs=float(k['Fs']),st=k['solver_status'],acc=k['accepted'],fb=k['fallback'],ds=k.get('dual_stat'),t0=float(r[1]))
        if ev=='solve_end' and sid in cur and cur[sid]['fb']=='True':
            k=kv(d); c=cur[sid]; c.update(hst=k.get('status'),obj=k.get('objective'),hsec=float(r[1])-c['t0'],el=el.get(sid)); out.append((sid,c))
    n5=[o for o in out if o[1]['st']=='5']
    print(p,'upper fallbacks',len(out),'st5',len(n5))
    for sid,c in out:
        o=float(c['obj']) if c['obj'] not in (None,'') else float('nan')
        flag='COUNTEREX' if c['st']=='5' and c['hst']=='OPTIMAL' and o>=1.000003 else ''
        print(f"  s{sid} el={c['el']} Fs={c['fs']:.6f} dllst={c['st']} ds={float(c['ds']):.4g} hsd={c['hst']} obj={c['obj']} sec={c['hsec']:.0f} {flag}")
    print('  st5 HSD seconds total',round(sum(c['hsec'] for s,c in n5)))
