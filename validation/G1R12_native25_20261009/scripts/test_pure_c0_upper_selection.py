"""Bounded upper witness selection for D03N on a separately saved D02A mesh.

This is an alternative mesh experiment, not replay of the native failed iterate.
"""
import json,time,hashlib
import numpy as np
import clarabel
from scipy import sparse
from prepare_inputs import ROOT
from case_locations import case_folder
from audit_batch import read_state
from model_support.upper_assembly import assemble
from common import save_mesh
from independent_audit import audit

def run(fs=1.20943742542898,experiment_tag=None):
    a=json.loads((ROOT/'inputs/D02A.json').read_text(encoding='utf-8'))
    b=json.loads((ROOT/'inputs/D03N.json').read_text(encoding='utf-8'))
    assert a['points']==b['points'] and a['faces']==b['faces'] and a['regions']==b['regions']
    native_path=case_folder('D02A')/'witness_upper.bin'
    m,_,meta=read_state(native_path,a['source_case']['policy'],'reference_total')
    m.materials=np.array([[p['c'],p['phi'],p['psi'],p['gamma']] for p in b['source_case']['materials']])
    m.provenance={'psi_policy':b['source_case']['policy']}
    m.physical_model_id='D03N_on_separate_D02A_native_mesh'
    assert np.all(m.materials[:,0]==0) and np.all(m.body[:,:,0]==0)
    q=8;p=assemble(m,fs,q)
    assert np.all(p.objective==0) and np.all(p.rhs[:-1]==0)
    n=len(p.objective);R=sparse.vstack([p.eq,-p.G],format='csc');rhs=np.r_[p.rhs,np.zeros(p.G.shape[0])]
    cones=[clarabel.ZeroConeT(len(p.rhs))]+[clarabel.SecondOrderConeT(d) if d==3 else clarabel.NonnegativeConeT(1) for d in p.dims]
    folder=ROOT/'pure_c0_upper_selection'
    if experiment_tag:folder=folder/experiment_tag
    folder.mkdir(parents=True,exist_ok=True)
    save_mesh(folder/'alternative_mesh.npz',m);records=[]
    for tag in ['constant_zero_objective','minimum_l2_qp','minimum_linf_socp']:
        matrix=R;bb=rhs;cc=p.objective;P=sparse.csc_matrix((n,n));kk=cones
        if tag=='minimum_l2_qp':P=sparse.eye(n,format='csc')
        elif tag=='minimum_linf_socp':
            R_aug=sparse.hstack([R,sparse.csc_matrix((R.shape[0],1))],format='csc')
            bound=sparse.hstack([sparse.vstack([sparse.eye(n),-sparse.eye(n)],format='csc'),-np.ones((2*n,1))],format='csc')
            matrix=sparse.vstack([R_aug,bound],format='csc');bb=np.r_[rhs,np.zeros(2*n)]
            cc=np.r_[p.objective,1.];P=sparse.csc_matrix((n+1,n+1));kk=cones+[clarabel.NonnegativeConeT(2*n)]
        opts=clarabel.DefaultSettings();opts.verbose=False;opts.max_iter=200;opts.time_limit=120.
        opts.tol_feas=1e-9;opts.tol_gap_abs=1e-9;opts.tol_gap_rel=1e-9
        start=time.perf_counter();sol=clarabel.DefaultSolver(P,cc,matrix,bb,kk,opts).solve()
        x=np.array(sol.x[:n]);before=audit(m,x,'upper',fs,q,0.,18*len(m.tri));work=before['work'][1]
        normalized=x/work if np.isfinite(work) and work>0 else x
        after=audit(m,normalized,'upper',fs,q,0.,18*len(m.tri))
        row={'selection':tag,'status':str(sol.status),'iterations':sol.iterations,
             'seconds':time.perf_counter()-start,'selection_objective':sol.obj_val,
             'max_abs_velocity_before':float(np.max(np.abs(x))),
             'max_abs_velocity_after':float(np.max(np.abs(normalized))),
             'original_physical_audit':before,'normalized_physical_audit':after,
             'upper_feasible_witness_pass':after['pass'] and after['value']<=1+1e-8}
        records.append(row);np.savez_compressed(folder/(tag+'.npz'),field=x,normalized_field=normalized,fs=fs,q=q)
        print(json.dumps(row,ensure_ascii=False),flush=True)
    out={'scope':'Python D03N physical-input experiment on a different native mesh; not native failure replay',
         'mesh_source_case':'D02A','mesh_source_sha256':hashlib.sha256(native_path.read_bytes()).hexdigest(),
         'mesh_metadata':meta,'physical_material_policy_case':'D03N','trial_fs':fs,
         'trial_origin':'rounded native log' if experiment_tag is None else 'selected positive control',
         'exact_native_failed_Fs_bits_available':False,'actual_native_failed_mesh_available':False,
         'vba_unchanged':True,'records':records,
         'minimum_linf_feasible_witness_confirmed':records[2]['upper_feasible_witness_pass']}
    (folder/'result.json').write_text(json.dumps(out,ensure_ascii=False,indent=2),encoding='utf-8')

if __name__=='__main__':run()
