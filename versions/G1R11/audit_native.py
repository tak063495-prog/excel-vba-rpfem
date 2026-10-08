from native_common import *
from common import load_mesh,physical_audit
from audit_precision import audit_precision
def field(path):
    raw=path.read_bytes();assert raw[:8]==b'P06X0001'
    n=struct.unpack_from('<i',raw,8)[0];x=np.frombuffer(raw,dtype='<f8',offset=12).copy();assert len(x)==n;return x
rows=[]
for mode in ['lower','upper']:
    p=R/'results'/f'P06N_{mode}.json'
    if not p.exists():continue
    rec=json.loads(p.read_text(encoding='utf-8'));assert rec['result'].startswith('SOLVED;'),rec
    m=load_mesh('P06N',path=R/'results/native1000_mesh.npz');x=field(R/'results'/f'P06N_{mode}.bin')
    au=physical_audit(m,x,mode,rec['fs'],rec['q']);assert au['pass'],au
    row={'test':rec['test'],'physical_audit':au,'scope':'FIXED_POINT_NOT_FULL_FS_ROOT'}
    if mode=='lower':
        assert x[-1]>=1;high=audit_precision(m,x/x[-1],rec['fs'],'lower');assert high['numerical_audit_pass'],high
        row['unit_load_50digit']=high
    else:assert au['value']<=1
    rows.append(row)
rootfile=R/'results/native_small_roots.json'
if rootfile.exists():
    rootrows=json.loads(rootfile.read_text(encoding='utf-8'))
    for rec in rootrows:
        assert rec['result'].startswith('PASS;'),rec
        prefix=Path(rec['prefix']);xy,tri,mat=read_geom(Path(str(prefix)+'_mesh.bin'))
        edges={}
        for t in tri:
            for a,b in zip(t,np.roll(t,-1)):
                k=tuple(sorted((int(a),int(b))));edges[k]=edges.get(k,0)+1
        bd={}
        for k,n in edges.items():
            if n==1:
                p=xy[list(k)];fixed=np.max(abs(p[:,1]))<1e-10 or np.max(abs(p[:,0]))<1e-10 or np.max(abs(p[:,0]-66))<1e-10
                bd[k]={'kind':'fixed' if fixed else 'load','rigid':-1,'loads':np.zeros((2,2))}
        body=np.zeros((len(tri),2,2));body[:,1,1]=-18.82
        m=Mesh(xy,tri,np.array([[5.,45.,45.,18.82]]),mat,np.full(len(tri),-1),body,bd,5.,'native_c5_small_associated',provenance={'case':'C5_small','psi_policy':'equal_reduced_phi'}).validate()
        for mode in ['lower','upper']:
            fs=float(rec['result'].split(mode+'=')[1].split(';')[0]);x=field(Path(str(prefix)+'_'+mode+'.bin'))
            au=physical_audit(m,x,mode,fs,2);assert au['pass'],au
            row={'test':'C5_small_'+rec['policy']+'_'+mode,'fs':fs,'physical_audit':au,'scope':'NATIVE_SMALL_ROOT_PHYSICAL_AUDIT_NOT_P06_ROOT'}
            if mode=='lower':
                high=audit_precision(m,x/x[-1],fs,'lower');assert high['numerical_audit_pass'],high;row['unit_load_50digit']=high
            rows.append(row)
    a=Path(rootrows[0]['prefix']);b=Path(rootrows[1]['prefix'])
    for kind in ['lower','upper','mesh']:
        assert Path(str(a)+'_'+kind+'.bin').read_bytes()==Path(str(b)+'_'+kind+'.bin').read_bytes()
    (R/'results/single_material_exact_match.json').write_text(json.dumps({'fields_and_mesh_bytes_identical':True,'lower_calls_identical':19,'upper_calls_identical':39,'scope':'C5_PHI_PSI45_SMALL_SLOPE_20_ELEMENTS'},indent=2),encoding='utf-8')
(R/'results/native_independent_audits.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
print('Independent physical/unit-load audits PASS:',[r['test'] for r in rows],flush=True)
