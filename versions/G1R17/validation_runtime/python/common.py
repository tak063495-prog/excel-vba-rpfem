"""Read supplied meshes without Triangle/Excel; never mutate source artifacts."""
from pathlib import Path
import sys, json, hashlib, copy
import numpy as np
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'vendor'))
from model_support.schema import Mesh,topology
from model_support.upper_assembly import assemble,strength,reduced_cohesion
from model_support.lower_assembly import assemble_lower
from independent_audit import audit
from zero_phi import assemble_zero,audit_zero
CASES={x['id']:x for x in json.loads((ROOT/'data/cases.json').read_text())}
REF=ROOT/'data/reference_results'
def model_key(case):
    c=copy.deepcopy(case)
    for k in ('id','title','h','area'):c.pop(k,None)
    c['formulation']='FELA_P2_ASSOC' if c['policy']=='equal_reduced_phi' else 'FELA_P2_DAVIS_EQUIVALENT'
    return hashlib.sha256(json.dumps(c,sort_keys=True,separators=(',',':')).encode()).hexdigest()
def load_mesh(case,level=0,path=None):
    c=CASES[case]; p=Path(path) if path else REF/f'{case}_M{level}_mesh.npz'
    with np.load(p,allow_pickle=False) as z:
        bounds={tuple(map(int,k)):{'kind':str(t),'rigid':-1,'loads':f.copy()} for k,t,f in zip(z['boundary_keys'],z['boundary_kind'],z['boundary_loads'])}
        mats=z['materials'].copy(); xy=z['xy'].copy();tri=z['tri'].copy()
        scale=float(z['stress_scale']) if 'stress_scale' in z else max(1.,max(mats[:,0]),max(mats[:,3])*np.ptp(xy[:,1]))
        m=Mesh(xy,tri,mats,z['material_id'].copy(),np.full(len(tri),-1),z['body'].copy(),bounds,scale,model_key(c),provenance={'psi_policy':c['policy'],'level':level,'case':case})
    return m.validate()
def save_mesh(path,m,**kwargs):
    np.savez_compressed(path,xy=m.xy,tri=m.tri,materials=m.materials,material_id=m.material_id,body=m.body,
        boundary_keys=np.array(list(m.boundary)),boundary_kind=np.array([v['kind'] for v in m.boundary.values()]),
        boundary_loads=np.array([v['loads'] for v in m.boundary.values()]),stress_scale=m.stress_scale,**kwargs)
def physical_audit(m,x,mode,fs,q=2):
    if mode=='lower':return audit(m,x,mode,fs,q,0.,18*len(m.tri))
    if np.all(m.materials[:,1]==0):
        pr=assemble_zero(m,fs,q); a=audit_zero(m,x,fs,q,float(pr.objective@x))
    else:
        pr=assemble(m,fs,q);a=audit(m,x,mode,fs,q,float(pr.objective@x),18*len(m.tri))
    return a

def dump(path,obj):
    def default(x):
        if isinstance(x,np.ndarray):return x.tolist()
        if isinstance(x,np.generic):return x.item()
        if isinstance(x,Path):return str(x)
        raise TypeError(type(x).__name__)
    Path(path).write_text(json.dumps(obj,indent=2,ensure_ascii=False,default=default),encoding='utf-8')
