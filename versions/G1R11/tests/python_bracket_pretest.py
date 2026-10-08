"""Only scheduling simulation. No FEM solution/status is inferred from these tests."""
from pathlib import Path
import math,random,json
R=Path(__file__).resolve().parents[1]
def solve(root,hint,cap=None,fail_cap=False):
    calls=[];used=False;failed=False
    def f(x):
        nonlocal failed
        calls.append(x)
        if fail_cap and cap is not None and x==cap and not failed:
            failed=True;raise ArithmeticError('KNOWN_NUMERICAL_FAILURE')
        return root/x-1.000001
    a,b=(hint*.9,hint*1.1) if hint>0 else (1.,2.)
    legacy_b=b
    if cap is not None:
        b=min(b,cap)
        if a>=b:a=b/2
    fa=f(a)
    try:fb=f(b)
    except ArithmeticError:
        cap=None;b=legacy_b;fb=f(b)
    while fa<0:
        b,fb=a,fa;a/=2;fa=f(a)
    while fb>0:
        a,fa=b,fb;next_b=b*2
        if cap is not None:
            if b>=cap:cap=None
            else:next_b=min(next_b,cap)
        try:fb=f(next_b)
        except ArithmeticError:
            cap=None;next_b=b*2;fb=f(next_b)
        b=next_b
    initial=(a,b)
    for _ in range(100):
        if b-a<=1e-4*max(1,(a+b)/2):break
        x=(a+b)/2;fx=f(x)
        if fx>=0:a,fa=x,fx
        else:b,fb=x,fx
    assert a<=root/1.000001<=b and fa>=0 and fb<=0
    return {'calls':calls,'bracket':(a,b),'initial':initial,'known_failure_fallback':failed}
rng=random.Random(20261007);rows=[]
for i in range(300):
    root=10**rng.uniform(-2,2);hint=root*rng.uniform(.4,1.4) if i%3 else 0
    cap=root*rng.uniform(1.001,1.8)
    rows.append({'legacy':solve(root,hint),'bounded':solve(root,hint,cap)})
for cap in [.4,.65]:solve(.66,.63,cap)
fail=solve(.66,0,.77,True);assert fail['known_failure_fallback']
demo={'legacy':solve(.66,.63),'bounded':solve(.66,.63,.77)}
(R/'results/python_bracket_pretest.json').write_text(json.dumps({'scope':'SIMULATED_MONOTONE_CALLBACK_NOT_FEM','random_cases_pass':300,'inconsistent_upper_hint_falls_back':True,'known_numeric_cap_failure_falls_back':True,'example':demo},indent=2),encoding='utf-8')
print('PASS 300 bracket scheduling simulations; sign checks; inconsistent/failed hint fallback')
