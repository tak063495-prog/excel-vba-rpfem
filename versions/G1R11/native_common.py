from pathlib import Path
import json,sys,hashlib,shutil,time,struct
import numpy as np
import win32com.client
R=Path(__file__).resolve().parent
sys.path.insert(0,str(R/'tests'));sys.path.insert(0,str(R))
from native_helpers import compile_book
from python_mesh_experiments import saved_mesh,refine,material_area,cross2,Mesh,CASES,model_key,save_mesh
identity=json.loads((R/'source_identity.json').read_text())
source=Path(identity['path'])
assert hashlib.sha256(source.read_bytes()).hexdigest()==identity['sha256']
NAMES=['RPX_Refine','RPX_Adapt','RPX_Fs']
def setvalue(book,key,value):
    ws=book.Worksheets('設定')
    for row in range(4,101):
        if ws.Cells(row,4).Value2==key:ws.Cells(row,2).Value2=value;return
    row={'ADAPT_REFINE_CAP':49,'FS_BRACKET_POLICY':50}[key]
    assert ws.Cells(row,4).Value2 is None
    ws.Cells(row,1).Value2={'ADAPT_REFINE_CAP':'追加細分化の要素上限','FS_BRACKET_POLICY':'Fs試行範囲の選び方'}[key]
    ws.Cells(row,4).Value2=key;ws.Cells(row,2).Value2=value
def inject(book,probe=True):
    for name in NAMES:
        cm=book.VBProject.VBComponents(name).CodeModule;cm.DeleteLines(1,cm.CountOfLines)
        cm.AddFromString('\r\n'.join((R/'src'/(name+'.bas')).read_text(encoding='utf-8').splitlines()[1:]))
    if probe:
        cm=book.VBProject.VBComponents.Add(1);cm.Name='MeshProbe'
        cm.CodeModule.AddFromString('\r\n'.join((R/'tests/MeshProbe.bas').read_text().splitlines()[1:]))
def write_input(m,path):
    edges={}
    for e,t in enumerate(m.tri):
        for a,b in zip(t,np.roll(t,-1)):edges.setdefault(tuple(sorted((int(a),int(b)))),[]).append(e)
    keys=[k for k,v in edges.items() if len(v)==1 or m.material_id[v[0]]!=m.material_id[v[1]]]
    path.write_bytes(struct.pack('<ii',len(m.xy),len(m.tri))+m.xy.astype('<f8').tobytes()+m.tri.astype('<i4').tobytes()+struct.pack('<i',len(keys))+np.array(keys,dtype='<i4').tobytes())
def read_geom(path):
    raw=path.read_bytes();nn,ne=struct.unpack_from('<ii',raw)
    xy=np.frombuffer(raw,dtype='<f8',offset=8,count=nn*2).reshape(nn,2).copy()
    tr=np.frombuffer(raw,dtype='<i4',offset=8+nn*16,count=ne*4).reshape(ne,4).copy()
    return xy,tr[:,:3],tr[:,3]
def canonical(xy,tri,mat):
    return sorted((int(k),tuple(sorted(tuple(float(v).hex() for v in xy[j]) for j in t))) for t,k in zip(tri,mat))
def geometry_mesh(xy,tri,mat):
    base=saved_mesh();edges={}
    for t in tri:
        for a,b in zip(t,np.roll(t,-1)):
            key=tuple(sorted((int(a),int(b))));edges[key]=edges.get(key,0)+1
    bd={}
    for k,n in edges.items():
        if n==1:
            p=xy[list(k)]
            kind='fixed' if np.max(abs(p[:,1]))<1e-10 else 'roller_x' if np.max(abs(p[:,0]))<1e-10 or np.max(abs(p[:,0]-28))<1e-10 else 'load'
            bd[k]={'kind':kind,'rigid':-1,'loads':np.zeros((2,2))}
    body=np.zeros((len(tri),2,2));body[:,1,1]=-19
    return Mesh(xy,tri,base.materials.copy(),mat,np.full(len(tri),-1),body,bd,20,base.physical_model_id,provenance={'case':'P06N','psi_policy':'hold_initial','source':'native_refine_legacy781_to_1000'}).validate()
