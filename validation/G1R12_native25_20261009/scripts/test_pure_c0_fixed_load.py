"""Python-only pure-c0 feasibility experiments on the actual native D01A mesh."""
from pathlib import Path
import json,time,hashlib
import numpy as np
import clarabel
from scipy import sparse
from prepare_inputs import ROOT
from case_locations import case_folder
from audit_batch import read_state
from model_support.lower_assembly import assemble_lower
from independent_audit import audit

def run():
    folder=case_folder('D01A')
    native=json.loads((folder/'status.json').read_text(encoding='utf-8'))
    rec=json.loads((ROOT/'inputs/D01A.json').read_text(encoding='utf-8'))
    path=folder/'witness_lower.bin'
    m,field,meta=read_state(path,rec['source_case']['policy'],'reference_total')
    assert np.all(m.materials[:,0]==0) and np.all(m.body[:,:,0]==0)
    assert all(np.all(e['loads'][:,0]==0) for e in m.boundary.values())
    records=[];out=ROOT/'pure_c0_feasibility';out.mkdir(exist_ok=True)
    for tag,fs,fixed in [('maximize_load_at_lower',native['lower_fs'],None),
                         ('unit_load_at_lower',native['lower_fs'],1.),
                         ('unit_load_at_upper',native['upper_fs'],1.)]:
        start=time.perf_counter();A,b,G,h,dims,zero=assemble_lower(m,fs,fixed_load=fixed)
        matrix=sparse.vstack([A,-G],format='csc');rhs=np.r_[b,h]
        q=np.zeros(A.shape[1])
        if fixed is None:q[-1]=-1.
        settings=clarabel.DefaultSettings();settings.verbose=False
        settings.tol_gap_abs=1e-9;settings.tol_gap_rel=1e-9;settings.tol_feas=1e-9
        settings.max_iter=200;settings.time_limit=120.
        sol=clarabel.DefaultSolver(sparse.csc_matrix((len(q),len(q))),q,matrix,rhs,
                                   [clarabel.ZeroConeT(len(b))]+[clarabel.SecondOrderConeT(d) for d in dims],settings).solve()
        x=np.asarray(sol.x);z=np.asarray(sol.z);status=str(sol.status)
        row={'tag':tag,'fs':fs,'fixed_load':fixed,'status':status,'iterations':sol.iterations,
             'seconds':time.perf_counter()-start,'variables':len(q),'equality_rows':len(b),
             'zero_stress_reduction_nodes':int(zero.sum()),'solver_primal_residual':sol.r_prim,'solver_dual_residual':sol.r_dual}
        if status in ['Solved','AlmostSolved','DualInfeasible'] and x[-1]>0:
            normalized=x/x[-1]
            au=audit(m,normalized,'lower',fs,meta['q'],-1.,18*len(m.tri))
            raw=max(max(0.,v)*m.stress_scale for v in au['errors'].values())
            row.update({'returned_load':float(x[-1]),'normalized_physical_audit':au,
                        'raw_physical_error':raw,'unit_witness_pass':au['pass'] and raw<=1e-7})
        if status=='PrimalInfeasible':
            negative_work=float(rhs@z)
            if negative_work<0:
                z=z/(-negative_work)
                residual=float(np.max(np.abs(matrix.T@z)))
                cone_error=0.;offset=len(b)
                for d in dims:
                    v=z[offset:offset+d];cone_error=max(cone_error,float(np.linalg.norm(v[1:])-v[0]));offset+=d
                row['infeasibility_certificate']={'normalized_rhs_dot':float(rhs@z),
                    'stationarity_max':residual,'dual_cone_error':max(0.,cone_error),
                    'numeric_pass':residual<=1e-8 and cone_error<=1e-8}
        np.savez_compressed(out/(tag+'.npz'),field=x,certificate=z,fs=fs)
        records.append(row);print(json.dumps(row,ensure_ascii=False),flush=True)
    result={'scope':'Python equivalent pure-c0 feasibility experiment; not an Excel root or 25-case PASS',
            'native_witness_sha256':hashlib.sha256(path.read_bytes()).hexdigest(),
            'native_mesh_metadata':meta,'same_materials_geometry_loads':True,
            'vba_unchanged':True,'records':records,
            'applicability':'verified scope: all cohesion zero, total reference loads; mixed-c P06 not validated and excluded from this proposed path',
            'fixed_unit_feasibility_confirmed':records[1].get('unit_witness_pass',False),
            'outside_infeasibility_confirmed':records[2].get('infeasibility_certificate',{}).get('numeric_pass',False)}
    (out/'result.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')

if __name__=='__main__':run()
