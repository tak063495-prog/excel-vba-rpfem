"""Python-only candidate mesh repair for D07A; does not patch VBA.

Preserve the complete traction step and all material laws. Split triangles in a
two-triangle fan at the traction discontinuity into centroid fans, then resolve
the same lower equations and audit the full solution independently.
"""
import json,sys
import numpy as np
import clarabel
from scipy import sparse
from prepare_inputs import ROOT,REPO
sys.path.insert(0,str(REPO/'versions/G1R12/validation_runtime/python'))
from common import Mesh,save_mesh,load_mesh
from model_support.lower_assembly import assemble_lower
from independent_audit import audit
from conic_lp import affine_reduce

def get_mesh():
    p=ROOT/'surcharge_diagnosis/mesh.npz'
    return load_mesh('D07A',path=p)

def fix(m):
    xy=m.xy.tolist();tri=m.tri.tolist();ids=m.material_id.tolist();changed=[]
    for point in [(2,10),(8,10)]:
        node=int(np.flatnonzero(np.all(m.xy==point,axis=1))[0])
        fan=np.flatnonzero(np.any(m.tri==node,axis=1))
        if len(fan)!=2:continue
        for e in fan:
            a,b,c=tri[e];center=len(xy);xy.append(np.mean(m.xy[[a,b,c]],axis=0).tolist())
            tri[e]=[a,b,center];tri.extend([[b,c,center],[c,a,center]]);ids.extend([ids[e],ids[e]]);changed.append(int(e))
    ids=np.array(ids);body=np.zeros((len(tri),2,2));body[:,1,1]=-m.materials[ids,3]
    out=Mesh(np.array(xy),np.array(tri),m.materials.copy(),ids,np.full(len(tri),-1),body,m.boundary.copy(),m.stress_scale,m.physical_model_id,provenance=m.provenance.copy()).validate()
    def a(z):
        u=z.xy[z.tri[:,1]]-z.xy[z.tri[:,0]];v=z.xy[z.tri[:,2]]-z.xy[z.tri[:,0]]
        return (u[:,0]*v[:,1]-u[:,1]*v[:,0]).sum()/2
    assert abs(a(out)-a(m))<1e-10
    return out,changed

def solve(m,fs,fixed_load=None):
    A,b,G,h,dims,zero=assemble_lower(m,fs,fixed_load=fixed_load)
    opt=clarabel.DefaultSettings();opt.verbose=False
    opt.tol_gap_abs=1e-9;opt.tol_gap_rel=1e-9;opt.tol_feas=1e-9
    q=np.zeros(A.shape[1])
    if fixed_load is None:q[-1]=-1.
    original_n=len(q);T=sparse.eye(original_n,format='csc');offset=np.zeros(original_n);reduction=[]
    if fixed_load is not None:
        A,b,G,h,dims,T,offset,reduction=affine_reduce(A,b,G,h,dims)
        q=np.asarray(q@T).ravel()
    sol=clarabel.DefaultSolver(sparse.csc_matrix((len(q),len(q))),q,sparse.vstack([A,-G],format='csc'),np.r_[b,h],[clarabel.ZeroConeT(A.shape[0])]+[clarabel.SecondOrderConeT(d) for d in dims],opt).solve()
    x=np.asarray(T@np.array(sol.x)+offset);au=audit(m,x,'lower',fs,4,-x[-1],18*len(m.tri))
    return x,{'status':str(sol.status),'iterations':sol.iterations,'load':float(x[-1]),'physical_audit':au,'fixed_load':fixed_load,'affine_reduction_variables':[original_n,len(q)],'scope':'feasible witness at fixed Fs, not a completed Fs root search'}

def run():
    folder=ROOT/'surcharge_diagnosis';m=get_mesh();fixed,changed=fix(m)
    out={'scope':'Python candidate only; baseline 25 native results unchanged','changed_triangles':changed,'original_elements':len(m.tri),'repaired_elements':len(fixed.tri),'fs':1.,'boundary_loads_unchanged':True,'materials_unchanged':True}
    for tag,model in [('original',m),('centroid_fan',fixed)]:
        x,row=solve(model,1.,fixed_load=1. if tag=='centroid_fan' else None);out[tag]=row
        np.savez_compressed(folder/(tag+'_lower_Fs1.npz'),field=x,fs=1.)
        save_mesh(folder/(tag+'_mesh.npz'),model)
    assert abs(out['original']['load'])<1e-7,out
    repaired=out['centroid_fan']
    repaired['optimization_converged']=repaired['status'] in ['Solved','AlmostSolved']
    repaired['feasibility_witness_pass']=repaired['physical_audit']['pass'] and abs(repaired['load']-1.)<1e-8
    assert repaired['feasibility_witness_pass'],out
    out['candidate_test_pass']=True
    out['candidate_test_scope']='Physical feasibility restored at Fs1/load1. Clarabel NumericalError is recorded; no optimization convergence or full Fs root PASS is claimed.'
    (folder/'candidate_repair.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')
    print(json.dumps(out,ensure_ascii=False),flush=True)

if __name__=='__main__':run()
