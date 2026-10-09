from pathlib import Path
import struct,json,sys
ROOT=Path(__file__).resolve().parents[1]
OLD=Path('C:/Users/link_/Desktop/RPFEM_G1R12_Native25_20261009')
DIAG=Path('C:/Users/link_/Desktop/RPFEM_D03N_Trace_20261009')
sys.path.insert(0,str(OLD))
from audit_batch import read_state
from audit_precision import audit_precision
from fractions import Fraction
import numpy as np
m,original,meta=read_state(DIAG/'capture/candidate_s59_c1_i50.mesh.bin','DAVIS_EQUIVALENT','reference_total');m.provenance['case']='D03N'
def field(p):
    raw=p.read_bytes();assert raw[:8]==b'P06X0001';n=struct.unpack_from('<i',raw,8)[0];assert len(raw)==12+8*n
    return np.frombuffer(raw,dtype='<f8',offset=12).copy()
def exact_work(x):
    acc=Fraction(0)
    for e,(a,b,c) in enumerate(m.tri):
        ax,ay=map(Fraction.from_float,m.xy[a]);bx,by=map(Fraction.from_float,m.xy[b]);cx,cy=map(Fraction.from_float,m.xy[c])
        area=((bx-ax)*(cy-ay)-(by-ay)*(cx-ax))/2
        for k in [3,4,5]:
            for j in [0,1]:acc+=area*Fraction.from_float(m.body[e,j,1])*Fraction.from_float(x[12*e+2*k+j])/(3*Fraction.from_float(m.stress_scale))
    return float(acc)
rec={'scope':'actual native exported fields; independent 50-digit audit plus exact stored-geometry/body-work integral','original':{'work':exact_work(original),'audit':audit_precision(m,original,meta['fs'],'upper')}}
for tag in ['original_normalized','solved']:
    p=ROOT/'single_point/D03N'/f'{tag}.x.bin'
    if p.exists():
        x=field(p);a=audit_precision(m,x,meta['fs'],'upper');rec[tag]={'work':exact_work(x),'audit':a}
        assert a['accepted'],rec[tag]
assert not rec['original']['audit']['accepted'],rec['original']
(ROOT/'results/work_oracle.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2),encoding='utf-8')
print(json.dumps({k: {'work':v['work'],'pass':v['audit']['accepted']} for k,v in rec.items() if isinstance(v,dict)},ensure_ascii=False))
