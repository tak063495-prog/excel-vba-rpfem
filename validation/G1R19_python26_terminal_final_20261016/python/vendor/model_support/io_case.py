from pathlib import Path
import struct,hashlib,json
import numpy as np
from scipy import sparse
from .schema import Mesh

def number(h):return struct.unpack('<d',bytes.fromhex(h))[0]

def parse_model(text,physical_id=None):
    rows=text.split('|');header=rows[0].split(',');n,e,m,r,ne=map(int,header[:5]);assert r==0,'RIGID_UNSUPPORTED'
    xy=np.array([[number(s[:16]),number(s[16:])]for s in rows[1:1+n]])
    tr=[];mi=[];ri=[];body=[]
    for s in rows[1+n:1+n+e]:
        v=s.split(',');tr.append(list(map(int,v[:3])));mi.append(int(v[3]));ri.append(int(v[4]));body.append([[number(t[:16]),number(t[16:])]for t in v[5:]])
    mat=np.array([[number(s[j*16:(j+1)*16])for j in range(4)]for s in rows[1+n+e:1+n+e+m]])
    boundary={}
    for s in rows[1+n+e+m:1+n+e+m+ne]:
        v=s.split(',')
        if int(v[3])<0:boundary[tuple(map(int,v[:2]))]={'kind':v[4],'rigid':int(v[5]),'loads':[[number(t[:16]),number(t[16:])]for t in v[6:]]}
    return Mesh(xy,np.array(tr,dtype=np.int64),mat,np.array(mi),np.array(ri),np.array(body),boundary,number(header[-1]),physical_id or hashlib.sha256(text.encode()).hexdigest(),header[-2]).validate()

def read_snapshot(path):
    path=Path(path);lines=path.read_text(encoding='utf-16').splitlines();rows=[s.split('\t')for s in lines[1:]]
    exact=next(v[1]for v in rows if v[0]=='exact_model');raw=next((v[1]for v in rows if v[0]=='raw_exact_model'),None)
    mesh=parse_model(exact,hashlib.sha256((raw or exact).encode()).hexdigest());vectors={};eq=[];co=[];cones=[];meta={'header':lines[0],'source_sha256':hashlib.sha256(path.read_bytes()).hexdigest(),'raw_model_available':bool(raw)}
    if raw:
        original=parse_model(raw)
        for key,value in mesh.boundary.items():
            source=original.boundary[key]
            value['free_in_raw']=source['kind']=='load'and source['rigid']<0 and not np.any(np.asarray(source['loads'])!=0)
    for v in rows:
        if v[0]in ['candidate_x','candidate_y','candidate_dual','q','eq_rhs','eq_row_norm']:
            vectors.setdefault(v[0],{})[int(v[1])]=number(v[2])
        elif v[0]=='Fs':meta['Fs']=number(v[1])
        elif v[0]=='cone':cones.append(list(map(int,v[1:])))
        elif v[0]=='eq':eq.append((int(v[1]),int(v[2]),number(v[3])))
        elif v[0]=='coef':co.append((int(v[1]),int(v[2]),int(v[3]),number(v[4])))
    vectors={k:np.array([v[i]for i in range(len(v))])for k,v in vectors.items()}
    n=len(vectors['q']);m=len(vectors['eq_rhs'])-1;first={c[0]:c[3]for c in cones};ns=sum(c[1]for c in cones)
    A=sparse.coo_matrix(([v[2]for v in eq],([v[0]for v in eq],[v[1]for v in eq])),shape=(m,n)).tocsc()
    G=sparse.coo_matrix(([v[3]for v in co],([first[v[0]]+v[1]for v in co],[v[2]for v in co])),shape=(ns,n)).tocsc()
    return mesh,meta,vectors,A,G,cones

def save_mesh(mesh,path,**arrays):
    path=Path(path);path.parent.mkdir(parents=True,exist_ok=True)
    data={'schema_version':1,'physical_model_id':mesh.physical_model_id,'mesh_id':mesh.mesh_id,'load_phase':mesh.load_phase,'stress_scale':mesh.stress_scale,'boundary':[{**v,'nodes':list(k),'loads':np.asarray(v['loads']).tolist()}for k,v in sorted(mesh.boundary.items())],'provenance':mesh.provenance}
    path.with_suffix('.json').write_text(json.dumps(data,indent=2),encoding='utf-8')
    np.savez_compressed(path.with_suffix('.npz'),xy=mesh.xy,tri=mesh.tri,materials=mesh.materials,material_id=mesh.material_id,rigid_id=mesh.rigid_id,body=mesh.body,**arrays)

def load_mesh(path):
    path=Path(path);d=json.loads(path.with_suffix('.json').read_text());a=np.load(path.with_suffix('.npz'))
    boundary={tuple(v['nodes']):{k:x for k,x in v.items()if k!='nodes'}for v in d['boundary']}
    m=Mesh(*(a[k].copy()for k in ['xy','tri','materials','material_id','rigid_id','body']),boundary,d['stress_scale'],d['physical_model_id'],d['load_phase'],d['provenance']).validate()
    assert m.mesh_id==d['mesh_id'],'SNAPSHOT_IDENTITY_CHANGED'
    return m

def write_native_model(mesh,path,fs,q=2):
    assert mesh.load_phase=='reference_total'
    def row(values):return ','.join('"'+v+'"'if isinstance(v,str)else format(float(v),'.17g')for v in values)
    lines=[row([len(mesh.xy),len(mesh.tri),len(mesh.materials),0,len(mesh.boundary),mesh.stress_scale,fs,q])]
    for c,phi,psi,gamma in mesh.materials:
        assert phi==psi
        lines.append(row([c,phi,gamma]))
    lines.extend(row(v)for v in mesh.xy)
    for i,t in enumerate(mesh.tri):lines.append(row([*t,mesh.material_id[i],mesh.rigid_id[i],*mesh.body[i,:,0],*mesh.body[i,:,1]]))
    for (a,b),v in mesh.boundary.items():
        loads=np.asarray(v['loads']);lines.append(row([a,b,v['kind'],v['rigid'],*loads[:,0],*loads[:,1]]))
    Path(path).write_text('\n'.join(lines)+'\n',encoding='ascii')
