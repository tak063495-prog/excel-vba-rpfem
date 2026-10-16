"""Direct original-cone Farkas feasibility for zero-cost upper programs only.

This diagnoses a discrete infeasibility sign. It never publishes a primal
upper bound or claims a continuum certificate. Original residuals, rather
than the auxiliary solver status, determine whether a sign is available.
"""
from __future__ import annotations
import argparse
import hashlib
import io
import json
from pathlib import Path
import time

import accuracy_driver as driver
from experiment_core import np,sparse,clarabel,Prepared,cone_error,write


def require_zero_cost(pr):
    if pr.mode!='upper' or pr.q!=8 or not np.array_equal(pr.original[-1],np.zeros_like(pr.original[-1])):
        raise ValueError('Only original all-zero-cost q8 upper programs are eligible')


def audit_farkas(pr,vector,target=1.000003):
    require_zero_cost(pr)
    E,b,G,h,dims,cost=pr.original
    if len(vector)!=len(b)+len(h) or not np.isfinite(vector).all():
        return {'accepted':False,'positive_discrete_sign':None,'reason':'NONFINITE_OR_WRONG_SIZE'}
    y,z=vector[:len(b)],vector[len(b):]
    stat=float(np.max(np.abs(E.T@y-G.T@z),initial=0.))
    cone=cone_error(z,dims);separation=-float(b@y+h@z)
    normal_error=abs(separation-target)
    accepted=bool(stat<=1e-7 and cone<=1e-8 and separation>=1.000001 and normal_error<=1e-7)
    return {'accepted':accepted,'positive_discrete_sign':1 if accepted else None,
        'dual_stationarity':stat,'dual_cone_error':cone,'dual_objective':separation,
        'normalization_error':normal_error,'target':target,
        'discrete_infeasibility_only':True,'primal_bound':False}


def direct_farkas(pr,regularization=0.,equilibrate=True,time_limit=120.,max_iter=350):
    require_zero_cost(pr);E,b,G,h,dims,cost=pr.original
    t=time.perf_counter();ny,nz=len(b),len(h);nv=ny+nz
    stationary=sparse.hstack([E.T,-G.T],format='csc')
    separation=sparse.csc_matrix(np.r_[-b,-h].reshape(1,-1))
    conic=sparse.hstack([sparse.csc_matrix((nz,ny)),-sparse.eye(nz,format='csc')],format='csc')
    A=sparse.vstack([stationary,separation,conic],format='csc')
    rhs=np.r_[np.zeros(len(cost)),1.000003,np.zeros(nz)]
    cones=[clarabel.ZeroConeT(len(cost)+1)]+[clarabel.NonnegativeConeT(d) if d==1 else clarabel.SecondOrderConeT(d) for d in dims]
    P=sparse.eye(nv,format='csc')*regularization if regularization else sparse.csc_matrix((nv,nv))
    settings=driver.core.settings({'equilibrate_enable':equilibrate,'time_limit':time_limit},max_iter)
    build=time.perf_counter()-t;t=time.perf_counter()
    solver=clarabel.DefaultSolver(P,np.zeros(nv),A,rhs,cones,settings);solver.solve();solution=solver.get_solution()
    vector=np.asarray(solution.x);audit=audit_farkas(pr,vector)
    audit.update(case=pr.m.provenance['case'],mode=pr.mode,q=pr.q,fs=pr.fs,
        mesh_id=pr.m.mesh_id,physical_model_id=pr.m.physical_model_id,elements=len(pr.m.tri),
        status=str(solution.status),iterations=solution.iterations,regularization=regularization,
        equilibrate_enable=equilibrate,build_seconds=build,solve_seconds=time.perf_counter()-t,
        original_equalities=ny,original_cone_components=nz,original_variables=len(cost))
    return audit,vector


def run(case,source,values,output,regularizations,equilibration,time_limit):
    output.mkdir(parents=True,exist_ok=True)
    recpath=driver.CATALOG/f'{case}.json';rec=json.loads(recpath.read_text(encoding='utf8'))
    if rec['source_case'].get('pore'):raise ValueError('Coupled pore snapshots are required')
    raw=(source/'best_upper_mesh.npz').read_bytes()
    with np.load(io.BytesIO(raw),allow_pickle=False) as data:
        mesh=driver.core.mesh_from_arrays(data['xy'].copy(),data['tri'].copy(),data['material_id'].copy(),rec)
    driver.save_mesh(output/'original_upper_mesh.npz',mesh)
    rows=[]
    for fs in values:
        pr=Prepared(mesh,fs,'upper',8,reduction='none');require_zero_cost(pr)
        for regularization in regularizations:
            for equilibrate in equilibration:
                row,vector=direct_farkas(pr,regularization,equilibrate,time_limit)
                rows.append(row);stem=f'fs_{fs:.15f}_reg_{regularization:g}_eq_{int(equilibrate)}'
                np.savez_compressed(output/(stem+'_dual.npz'),dual=vector,fs=fs,mode='upper',q=8)
                write(output/(stem+'_audit.json'),row)
                write(output/'diagnosis.json',{'case':case,'source':str(source),'source_mesh_sha256':hashlib.sha256(raw).hexdigest(),
                    'catalog_sha256':hashlib.sha256(recpath.read_bytes()).hexdigest(),'trials':rows,
                    'references':['https://clarabel.org/stable/python/getting_started_py/',
                                  'https://docs.scipy.org/doc/scipy/reference/generated/scipy.sparse.linalg.lsmr.html']})
                print('FARKAS',case,fs,'reg',regularization,'eq',equilibrate,json.dumps(row),flush=True)
                if row['accepted']:break
            if rows[-1]['accepted']:break
    return rows


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--case',default='D01A');parser.add_argument('--source',type=Path,required=True)
    parser.add_argument('--fs',type=float,nargs='+',required=True);parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--regularization',type=float,nargs='+',default=[0.,1e-10])
    parser.add_argument('--equilibration',choices=['both','on','off'],default='both')
    parser.add_argument('--time-limit',type=float,default=120.)
    args=parser.parse_args()
    eq=[True,False] if args.equilibration=='both' else [args.equilibration=='on']
    run(args.case,args.source,args.fs,args.output,args.regularization,eq,args.time_limit)
