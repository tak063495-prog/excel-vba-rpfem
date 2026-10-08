from pathlib import Path
import sys,json,struct
import numpy as np
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'validation_runtime/python'))
from common import Mesh, physical_audit, dump
from audit_precision import audit_precision
def field(path):
    raw=path.read_bytes();assert raw[:8]==b'P06X0001'
    n=struct.unpack_from('<i',raw,8)[0];x=np.frombuffer(raw,dtype='<f8',offset=12).copy();assert len(x)==n;return x
def mesh(path):
    raw=path.read_bytes();nn,ne=struct.unpack_from('<ii',raw)
    xy=np.frombuffer(raw,dtype='<f8',offset=8,count=nn*2).reshape(nn,2).copy()
    tr=np.frombuffer(raw,dtype='<i4',offset=8+nn*16,count=ne*4).reshape(ne,4).copy()
    tri,mat=tr[:,:3],tr[:,3];edges={}
    for t in tri:
        for a,b in zip(t,np.roll(t,-1)):
            k=tuple(sorted((int(a),int(b))));edges[k]=edges.get(k,0)+1
    bd={}
    for k,n in edges.items():
        if n==1:
            p=xy[list(k)];fixed=np.max(abs(p[:,1]))<1e-10 or np.max(abs(p[:,0]))<1e-10 or np.max(abs(p[:,0]-66))<1e-10
            bd[k]={'kind':'fixed' if fixed else 'load','rigid':-1,'loads':np.zeros((2,2))}
    body=np.zeros((ne,2,2));body[:,1,1]=-18.82
    return Mesh(xy,tri,np.array([[5.,45.,45.,18.82]]),mat,np.full(ne,-1),body,bd,5.,'native_c5_small',provenance={'case':'C5_small','psi_policy':'equal_reduced_phi'}).validate()
def run():
    rows=[];recs={}
    for version in ['G1R11','G1R12']:
        rec=json.loads((ROOT/'results'/f'native_{version}.json').read_text(encoding='utf-8'));assert rec['pass']
        recs[version]=rec;exports=ROOT/Path(rec['exports'])
        for mode in ['lower','upper']:
            fs=float(rec['roots'].split(mode+'=')[1].split(';')[0]);x=field(exports/f'roots_{mode}.bin')
            m=mesh(exports/'roots_mesh.bin');au=physical_audit(m,x,mode,fs,2);assert au['pass'],au
            row={'test':version+'_root_'+mode,'fs':fs,'physical_audit':au}
            if mode=='lower':
                hi=audit_precision(m,x/x[-1],fs,'lower');assert hi['numerical_audit_pass'];row['unit_load_50digit']=hi
            rows.append(row)
        for scenario in ['normal']+(['factor'] if version=='G1R12' else []):
            m=mesh(exports/f'{scenario}_mesh.bin');x=field(exports/f'{scenario}.bin')
            au=physical_audit(m,x,'upper',3.,2);assert au['pass'],au
            rows.append({'test':version+'_'+scenario+'_fixed_upper_Fs3','physical_audit':au})
    a=ROOT/Path(recs['G1R11']['exports']);b=ROOT/Path(recs['G1R12']['exports'])
    comparisons={}
    for name in ['roots_lower.bin','roots_upper.bin','roots_mesh.bin','normal.bin','normal_mesh.bin']:
        pa=a/name;pb=b/name
        comparisons[name]={'bytes_identical':pa.read_bytes()==pb.read_bytes()}
        if name=='roots_mesh.bin' or name=='normal_mesh.bin':assert comparisons[name]['bytes_identical']
        else:
            delta=float(np.max(abs(field(pa)-field(pb))));comparisons[name]['max_field_delta']=delta
            assert delta<1e-7,(name,delta)
    dump(ROOT/'results/independent_audits.json',{'pass':True,'scope':'independent Python raw-equation audits of native fields; not full P06N/25 cases','results':rows,'old_new_comparisons':comparisons})
    print('Independent physical audits PASS',len(rows),comparisons)
if __name__=='__main__':run()
