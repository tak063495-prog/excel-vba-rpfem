"""Read-only D03N evidence and exact Python arithmetic micro-experiments.

These synthetic experiments do not replay the native failure operands and do
not claim to solve D03N. Fractions provide an oracle, not a fast VBA kernel.
"""
from pathlib import Path
from fractions import Fraction
import math, json, csv, hashlib
from prepare_inputs import ROOT, REPO
from case_locations import case_folder

def exact_dot(pairs):
    return sum((Fraction.from_float(a)*Fraction.from_float(b)
                for a,b in pairs), Fraction())

def scaled_dot(pairs):
    """One shared exponent: useful candidate, unsafe under deep cancellation."""
    terms=[]
    for a,b in pairs:
        if not (math.isfinite(a) and math.isfinite(b)):
            raise ValueError('nonfinite input')
        if a == 0 or b == 0: continue
        ma,ea=math.frexp(a); mb,eb=math.frexp(b)
        terms.append((ma*mb,ea+eb))
    if not terms:return 0.0
    exponent=max(e for _,e in terms)
    return math.ldexp(math.fsum(math.ldexp(m,e-exponent)
                               for m,e in terms),exponent)

def run():
    folder=case_folder('D03N')
    status=json.loads((folder/'status.json').read_text(encoding='utf-8'))
    logs=list((folder/'RPFEM_logs').glob('*analysis.tsv'));assert len(logs)==1
    with logs[0].open(encoding='utf-16') as f:
        rows=list(csv.reader(f,delimiter='\t'))
    failures=[r for r in rows if 'HSD_DD_PRODUCT_UNDERFLOW' in '\t'.join(r)]
    assert failures and status['error']=='HSD_DD_PRODUCT_UNDERFLOW'
    context=[r for r in rows[-500:] if len(r)>6 and r[6] not in ['stage','hopt0','hopt0_gauge','hopt0_coverage','run_phase']]
    source=REPO/'versions/G1R12/src/RPX_HSDPrecision.bas'
    guards=[{'line':i,'code':s.strip()} for i,s in enumerate(source.read_text(encoding='utf-8-sig').splitlines(),1) if 'HSD_DD_PRODUCT_UNDERFLOW' in s]
    tiny=math.ldexp(1.0,-538)
    cases=[
        ('ordinary_product',[(2.0,3.0)]),
        ('real_zero',[(0.0,tiny)]),
        ('positive_product_underflow',[(tiny,tiny)]),
        ('negative_product_underflow',[(-tiny,tiny)]),
        ('four_tiny_products_form_subnormal',[(tiny,tiny)]*4),
        ('cancel_large_keep_subnormal',[(1.0,1.0),(-1.0,1.0)]+[(tiny,tiny)]*4),
        ('cancel_large_keep_normal',[(1e200,1e100),(-1e200,1e100),(1e-100,1.0)]),
        ('mixed_scale_finite',[(1e290,1e-290),(1e-290,1e290)]),
    ]
    experiments=[]
    for name,pairs in cases:
        exact=exact_dot(pairs); expected=float(exact)
        naive=math.fsum(a*b for a,b in pairs)
        scaled=scaled_dot(pairs)
        guard=any(a!=0 and b!=0 and a*b==0 for a,b in pairs)
        experiments.append({'name':name,'pairs_hex':[[a.hex(),b.hex()] for a,b in pairs],
                            'native_guard_would_trigger':guard,
                            'exact_sign':(exact>0)-(exact<0),'correctly_rounded_hex':expected.hex(),
                            'ordinary_dot_hex':naive.hex(),'shared_exponent_dot_hex':scaled.hex(),
                            'ordinary_matches':naive==expected,'shared_exponent_matches':scaled==expected})
    byname={r['name']:r for r in experiments}
    assert byname['positive_product_underflow']['native_guard_would_trigger']
    assert byname['positive_product_underflow']['exact_sign']==1
    assert not byname['four_tiny_products_form_subnormal']['ordinary_matches']
    assert byname['four_tiny_products_form_subnormal']['shared_exponent_matches']
    assert not byname['cancel_large_keep_subnormal']['shared_exponent_matches']
    assert not byname['cancel_large_keep_normal']['shared_exponent_matches']
    assert all(r['ordinary_matches'] for r in experiments if r['name'] in ['ordinary_product','real_zero','mixed_scale_finite'])
    out=ROOT/'underflow_diagnosis';out.mkdir(exist_ok=True)
    result={'case':'D03N','native_failure_confirmed':True,'native_failure_replayed':False,
            'source_unchanged':True,'source_sha256':hashlib.sha256(source.read_bytes()).hexdigest(),
            'actual_operands_available':False,'guard_sites':guards,
            'native_log':str(logs[0]),'failure_rows':failures,'last_context':context[-45:],
            'experiments':experiments,'experiment_assertions_pass':True,
            'vba_repair_implemented':False,'native_case_pass':False}
    (out/'result.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'case':'D03N','experiment_assertions_pass':True,'synthetic_tests':len(experiments),
                      'actual_operands_available':False,'vba_repair_implemented':False}))

if __name__=='__main__':run()
