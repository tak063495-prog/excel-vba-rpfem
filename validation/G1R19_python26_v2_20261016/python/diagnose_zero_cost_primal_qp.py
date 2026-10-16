"""Minimum-velocity-norm field search under unchanged original zero-cost constraints."""
from pathlib import Path
import time
import argparse
import experiment_core as core
import accuracy_driver as driver
np,sparse,clarabel=core.np,core.sparse,core.clarabel


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--ridges',nargs='+',type=float,default=[1.,1e-4])
    parser.add_argument('--scaled',action='store_true')
    parser.add_argument('--strict',action='store_true')
    args=parser.parse_args()
    root=Path(__file__).resolve().parents[1]
    case,fs='D02A',1.50600141054559
    _,rec=driver.load_model(case)
    path=root/'results/d02_final/D02A/D02A_C7_mesh.npz'
    mesh=driver.load_saved_mesh(path,rec)
    pr=core.Prepared(mesh,fs,'upper',8,reduction='none')
    E,b,G,h,dims,cost=pr.original
    assert pr.mode=='upper' and pr.q==8 and np.array_equal(cost,np.zeros_like(cost))
    assert (pr.A-sparse.vstack([E,-G],format='csc')).nnz==0
    assert np.array_equal(pr.rhs,np.r_[b,h])
    row_scale=np.ones(pr.A.shape[0])
    if args.scaled:
        start=0
        for dimension in dims:
            block=G[start:start+dimension];norm=float(np.sqrt(block.multiply(block).sum()))
            if norm>1e-10:row_scale[len(b)+start:len(b)+start+dimension]=1/norm
            start+=dimension
    AA=(sparse.diags(row_scale)@pr.A).tocsc();bb=row_scale*pr.rhs
    output=root/'results'/('primal_qp_diagnostic_'+str(time.time_ns()))
    output.mkdir()
    report={'case':case,'fs':fs,'q':8,'mesh_id':mesh.mesh_id,'physical_model_id':mesh.physical_model_id,
        'mesh_path':str(path),'original_cost_exact_zero':True,'QP_dual_used':False,
        'cone_block_common_positive_scaling':args.scaled,'strict':args.strict,'runs':[]}
    for ridge in args.ridges:
        for equilibrate in (False,True):
            overrides={'equilibrate_enable':equilibrate,'time_limit':60.,'iterative_refinement_max_iter':30,
                'iterative_refinement_abstol':1e-15,'iterative_refinement_reltol':1e-15}
            if args.strict:overrides.update(tol_gap_abs=1e-13,tol_gap_rel=1e-13,tol_feas=1e-13)
            started=time.perf_counter()
            solver=clarabel.DefaultSolver(ridge*sparse.eye(len(cost),format='csc'),np.zeros_like(cost),
                AA,bb,pr.cones,core.settings(overrides,350))
            solution=solver.solve();x=np.asarray(solution.x)
            audit=core.evaluate(pr,x,None)
            record={'ridge':ridge,'equilibrate':equilibrate,'status':str(solution.status),
                'iterations':solution.iterations,'seconds':time.perf_counter()-started,
                'QP_objective':float(solution.obj_val),'field_norm':float(np.linalg.norm(x)),
                'original_audit':audit,'witness_only':True,'QP_dual_used':False}
            report['runs'].append(record)
            if audit['witness']:
                name=f'{case}_ridge{ridge}_eq{int(equilibrate)}_field.npz'
                np.savez_compressed(output/name,x=x,fs=fs,q=8,mode='upper')
                driver.save_mesh(output/name.replace('_field','_mesh'),mesh)
                record['field_path']=name
            core.write(output/'audit.json',report)
            print(ridge,equilibrate,record['status'],audit['physical_pass'],audit['witness'],
                audit['eq_error'],audit['cone_error'],flush=True)
    print('DIAGNOSTIC',output,flush=True)


if __name__=='__main__':main()
