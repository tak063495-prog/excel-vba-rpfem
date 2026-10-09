"""Mesh-only native export plus Python lower-load feasibility for D07A.

Check algebraic traction compatibility at surcharge endpoints before blaming
the optimizer. Uses a fresh private Excel copy and preserves every active run.
"""
import json, shutil, sys, math
import numpy as np
import clarabel
from scipy import sparse
import win32com.client
from prepare_inputs import ROOT, SOURCE, REPO
from mesh_identity import read
sys.path.insert(0,str(REPO/'versions/G1R12/tests'))
from native_helpers import compile_book
sys.path.insert(0,str(REPO/'versions/G1R12/validation_runtime/python'))
from common import Mesh, save_mesh
from model_support.lower_assembly import assemble_lower

def on(p,a,b):
    a=np.asarray(a);b=np.asarray(b);d=b-a;v=p-a;t=v@d/(d@d)
    return -1e-10<=t<=1+1e-10 and abs(d[0]*v[1]-d[1]*v[0])<1e-8

def run():
    folder=ROOT/'surcharge_diagnosis';folder.mkdir(exist_ok=True)
    path=folder/'D07A_mesh_only_target538.xlsm';assert not path.exists();shutil.copy2(SOURCE,path)
    rec=json.loads((ROOT/'inputs/D07A.json').read_text(encoding='utf-8'))
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;book=None
    try:
        book=app.Workbooks.Open(str(path),UpdateLinks=0)
        cm=book.VBProject.VBComponents.Add(1);cm.Name='MeshIdentity'
        cm.CodeModule.AddFromString('\r\n'.join((ROOT/'tests/MeshIdentity.bas').read_text(encoding='utf-8').splitlines()[1:]))
        compile_book(app,book)
        ws=book.Worksheets('要素定義');ws.Range('A7:T1006').ClearContents()
        for first,last,name in [('A','C','points'),('E','H','regions'),('J','T','faces')]:
            data=tuple(tuple(v) for v in rec[name]);rng=ws.Range(first+'7:'+last+str(6+len(data)));rng.Value2=data;assert rng.Value2==data
        ws=book.Worksheets('材料データ');ws.Range('A5:Q1004').ClearContents();ws.Range('A5:E5').Value2=tuple(tuple(v) for v in rec['materials'])
        result=app.Run("'"+book.Name+"'!MeshIdentityExport",str(folder/'mesh.bin'),538);assert result.startswith('OK '),result
    finally:
        if book is not None:book.Close(SaveChanges=False)
        app.Quit()
    xy,tr=read(folder/'mesh.bin');counts={}
    for t in tr[:,:3]:
        for a,b in zip(t,np.roll(t,-1)):
            key=tuple(sorted((int(a),int(b))));counts[key]=counts.get(key,0)+1
    kinds={'自由':'load','固定XY':'fixed','固定X':'roller_x','固定Y':'roller_y'};bd={}
    for key,count in counts.items():
        if count!=1:continue
        faces=[f for f in rec['boundary_explicit'] if all(on(p,f['a'],f['b']) for p in xy[list(key)])];assert len(faces)==1
        f=faces[0];bd[key]={'kind':kinds[f['kind']],'rigid':-1,'loads':np.column_stack([np.zeros(2),np.asarray(f['load0'])+f['load1']])}
    mats=np.asarray([[m['c'],m['phi'],m['psi'],m['gamma']] for m in rec['source_case']['materials']]);body=np.zeros((len(tr),2,2));body[:,1,1]=-mats[tr[:,3],3]
    scale=max(1.,float(mats[:,0].max()))
    m=Mesh(xy,tr[:,:3],mats,tr[:,3],np.full(len(tr),-1),body,bd,scale,'native_D07A_coarse',provenance={'psi_policy':'equal_reduced_phi'}).validate()
    save_mesh(folder/'mesh.npz',m)
    ends=[]
    for point in [(2,10),(8,10)]:
        idx=np.flatnonzero(np.all(xy==point,axis=1));assert len(idx)==1;idx=int(idx[0])
        tri_ids=np.flatnonzero(np.any(m.tri==idx,axis=1));around=m.tri[tri_ids]
        ends.append({'point':point,'node':idx,'triangles':tri_ids.tolist(),'fan_size':len(tri_ids),'triangle_coordinates':xy[around].tolist()})
    A,b,G,h,cones,zero=assemble_lower(m,1.,fixed_load=1.)
    opt=clarabel.DefaultSettings();opt.verbose=False
    sol=clarabel.DefaultSolver(sparse.csc_matrix((A.shape[1],A.shape[1])),np.zeros(A.shape[1]),A,b,[clarabel.ZeroConeT(A.shape[0])],opt).solve()
    z=np.asarray(sol.z);b_dot=float(b@z)
    if b_dot<0:z=z/(-b_dot)
    certificate={'status':str(sol.status),'iterations':sol.iterations,'normalized_A_transpose_z_inf':float(np.max(abs(A.T@z))),'normalized_b_dot_z':float(b@z)}
    certificate['numerical_infeasibility_certificate']=str(sol.status)=='PrimalInfeasible' and certificate['normalized_A_transpose_z_inf']<1e-7 and certificate['normalized_b_dot_z']<-.99
    out={'scope':'native coarse target538 from actual adapt_engine log; equality-only lower load=1 feasibility, independent of material strength','native_target':538,'stress_scale':scale,'source_materials':mats.tolist(),'nodes':len(xy),'elements':len(tr),'surcharge_resultant':-120.,'endpoints':ends,'linear_certificate':certificate,
         'interpretation':'If equality constraints alone exclude load=1, changing HSD tolerances or material strength cannot repair this mesh/traction representation.'}
    (folder/'result.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps({'nodes':len(xy),'elements':len(tr),'endpoint_fan_sizes':[r['fan_size'] for r in ends],'linear_certificate':certificate},ensure_ascii=False),flush=True)

if __name__=='__main__':run()
