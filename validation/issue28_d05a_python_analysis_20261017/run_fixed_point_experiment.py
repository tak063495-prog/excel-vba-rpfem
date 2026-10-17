"""Bounded, original-q8 D05A/P02A conditioning experiment; no VBA claim.

Each child solves one upper LP at Fs=1 on the recorded 214-element coarse
geometry. This is not a root search, 2% completion test, or Fix4 DLL replay.
"""
from __future__ import annotations
import os
for key in ('OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS','NUMEXPR_NUM_THREADS'):
    os.environ[key]='1'
import argparse
import datetime as dt
import hashlib
import json
from pathlib import Path
import platform
import struct
import subprocess
import sys
import time
import traceback

sys.dont_write_bytecode=True
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'python_snapshot'))
import experiment_core as core

def utc():return dt.datetime.now(dt.timezone.utc).isoformat()
def sha(raw):return hashlib.sha256(raw).hexdigest()
def write(path,value):core.write(path,value)

def frozen_hashes():
    expected=json.loads((ROOT/'SOURCE_DATA_MANIFEST.json').read_text(encoding='utf-8'))['source_and_data_files']
    actual={name:sha((ROOT/name).read_bytes()) for name in expected}
    for name,record in expected.items():
        if actual[name]!=record['sha256']:raise ValueError('Frozen source/data changed: '+name)
    actual['run_fixed_point_experiment.py']=sha(Path(__file__).read_bytes())
    return actual

def load_model(case):
    record=json.loads((ROOT/'inputs'/f'{case}.json').read_text(encoding='utf-8'))
    raw=(ROOT/'inputs'/f'{case}_coarse_geometry.bin').read_bytes()
    nn,ne=struct.unpack_from('<ii',raw)
    assert len(raw)==8+16*nn+16*ne
    xy=core.np.frombuffer(raw,'<f8',2*nn,8).copy().reshape(nn,2)
    cells=core.np.frombuffer(raw,'<i4',4*ne,8+16*nn).copy().reshape(ne,4)
    tri,ids=cells[:,:3],cells[:,3]
    points=xy[tri];a=points[:,1]-points[:,0];b=points[:,2]-points[:,0]
    area=(a[:,0]*b[:,1]-a[:,1]*b[:,0])/2
    actual=core.np.bincount(ids,weights=area,minlength=len(record['region_areas']))
    expected=core.np.asarray(record['region_areas'])
    area_error=float(core.np.max(abs(actual-expected)))
    # This only protects summation when L is 1000x. It does not change the
    # coordinates/operators: the original builder has an absolute 1e-9 area
    # assertion that is too small at a 158,000,000-unit area.
    bound=8*ne*core.np.finfo(float).eps*max(1.,float(core.np.sum(abs(area))))
    assert core.np.all(area>0) and area_error<=max(1e-9,bound)
    construction=dict(record,region_areas=actual.tolist())
    mesh=core.mesh_from_arrays(xy,tri,ids,construction)
    mesh.provenance['diagnostic_scope']='recorded 214-element coarse topology; fixed upper Fs=1'
    mesh.provenance['canonical_area_error']=area_error
    mesh.provenance['canonical_area_roundoff_bound']=bound
    return mesh,record

def program_digest(program):
    h=hashlib.sha256()
    for value in program:
        arrays=(value.indptr,value.indices,value.data) if core.sparse.issparse(value) else (core.np.asarray(value),)
        for array in arrays:
            h.update(str(array.shape).encode('ascii'));h.update(str(array.dtype).encode('ascii'));h.update(array.tobytes())
    return h.hexdigest()

def dual_metrics(pr,z):
    E,b,G,h,dims,cost=pr.original
    y,zc=z[:len(b)],z[len(b):]
    stat=float(core.np.max(abs(cost+E.T@y-G.T@zc),initial=0.))
    cone=core.cone_error(zc,dims)
    objective=-float(b@y+h@zc)
    return {'stationarity':stat,'cone_error':cone,'objective':objective,
        'original_upper_inverse_guide_accepted':bool(stat<=1e-7 and cone<=1e-8 and objective>=1.000003),
        'thresholds':{'raw_stationarity':1e-7,'dual_cone':1e-8,'original_dual_objective_min':1.000003},
        'scope':'discrete upper root sign guide only; not a continuum primal bound'}

def direct_transformed_solve(pr,route,seconds):
    E,b,G,h,dims,cost=pr.original
    qscale=max(1.,float(core.np.max(abs(cost),initial=0.)))
    rowscale=core.np.ones(pr.A.shape[0]);coordinate_scale=1.;objective_scale=1.
    if route=='normalized_objective':objective_scale=1./qscale
    elif route=='velocity_units':
        work=core.np.flatnonzero(b)
        assert len(work)==1 and b[work[0]]>0 and core.np.all(h==0)
        coordinate_scale=float(b[work[0]])
        rowscale[len(b):]=core.cone_row_scales(G,dims)
    else:raise ValueError(route)
    A=(core.sparse.diags(rowscale)@pr.A).tocsc()
    rhs=rowscale*pr.rhs/coordinate_scale
    transformed_cost=cost*objective_scale*coordinate_scale
    config={'time_limit':seconds}
    if route=='velocity_units':
        config.update(tol_gap_abs=1e-13,tol_gap_rel=1e-13,tol_feas=1e-13,
            iterative_refinement_max_iter=30,iterative_refinement_abstol=1e-15,iterative_refinement_reltol=1e-15)
    solver=core.clarabel.DefaultSolver(core.sparse.csc_matrix((len(cost),len(cost))),transformed_cost,A,rhs,pr.cones,core.settings(config,150))
    t=time.perf_counter();solution=solver.solve();elapsed=time.perf_counter()-t
    transformed_x=core.np.asarray(solution.x);transformed_z=core.np.asarray(solution.z)
    x=coordinate_scale*transformed_x
    # c' = objective_scale*coordinate_scale*c and rhs'=R*rhs/coordinate_scale.
    # The original dual is R*z'/(objective_scale*coordinate_scale).
    z=rowscale*transformed_z/(objective_scale*coordinate_scale)
    audit=core.evaluate(pr,x,z)
    result={'status':str(solution.status),'iterations':solution.iterations,'solver_seconds':elapsed,
        'backend_solve_seconds':float(solution.solve_time),'audit':audit,
        'transform':{'route':route,'objective_scale':objective_scale,'coordinate_scale':coordinate_scale,
            'common_positive_cone_row_scaling':route=='velocity_units',
            'transformed_reported_primal_objective':float(solution.obj_val),
            'original_primal_objective':float(cost@x),
            'restore_original_objective':float(solution.obj_val)/objective_scale,
            'objective_restore_absolute_error':abs(float(cost@x)-float(solution.obj_val)/objective_scale)},
        'original_dual':dual_metrics(pr,z),'incorrect_unrestored_dual_diagnostic':dual_metrics(pr,transformed_z),
        'scope':'Exactly restored original full program; transformed solver status alone supplies no sign or physical certificate'}
    return result,x,z,{'transformed_x':transformed_x,'transformed_z':transformed_z,'row_scale':rowscale}

def child(args):
    out=args.output.resolve()
    if out.exists():raise FileExistsError(out)
    out.mkdir(parents=True)
    before=frozen_hashes();started=time.perf_counter()
    record={'status':'RUNNING','case':args.case,'route':args.route,'fs':1.,'q':8,'pid':os.getpid(),
        'started_utc':utc(),'source_data_hashes_before':before,'python':sys.version,'platform':platform.platform(),
        'numpy_version':core.np.__version__,'scipy_version':__import__('scipy').__version__,'clarabel_version':core.clarabel.__version__,
        'solver_threads':1,'solve_time_limit_seconds':args.seconds,'optimizer_calls':0,
        'scope':'Fresh fixed-point upper LP diagnostic. No root search or GAP2% completion claim; not Fix4 HSD/native execution.'}
    write(out/'result.json',record)
    try:
        mesh,physical=load_model(args.case)
        pr=core.Prepared(mesh,1.,'upper',8,reduction='affine' if args.route=='affine' else 'none')
        original_hash=program_digest(pr.original)
        E,b,G,h,dims,cost=pr.original
        qscale=max(1.,float(core.np.max(abs(cost),initial=0.)))
        zero=core.np.zeros(len(b)+len(h))
        record['model']={'nodes':len(mesh.xy),'elements':len(mesh.tri),'mesh_id':mesh.mesh_id,
            'physical_model_id':mesh.physical_model_id,'policy':mesh.provenance['psi_policy'],
            'canonical_area_error':mesh.provenance['canonical_area_error'],'body_force_gamma':float(mesh.materials[0,3])}
        record['original_program']={'sha256':original_hash,'variables':len(cost),'equalities':len(b),
            'cone_rows':len(h),'cone_blocks':len(dims),'matrix_nonzeros':E.nnz+G.nnz,
            'max_abs_original_cost':float(core.np.max(abs(cost),initial=0.)),
            'VBA_style_QScale':qscale,'max_abs_normalized_cost':float(core.np.max(abs(cost/qscale),initial=0.)),
            'zero_original_dual':dual_metrics(pr,zero),
            'zero_candidate_normalized_objective_stationarity':float(core.np.max(abs(cost/qscale),initial=0.)),
            'zero_candidate_passes_relative_stationarity_if_1e7_times_cost':bool(core.np.max(abs(cost),initial=0.)<=1e-7*qscale)}
        record['optimizer_calls']=1
        if args.route in ('none','affine'):
            r,x,reduced_dual=core.solve(pr,{'time_limit':args.seconds},max_iter=150)
            # The unchanged experiment_core API returns canonical *reduced*
            # duals as its third value, even though its reported audit uses a
            # restored original dual. Restore it explicitly for saved vectors.
            z,restoration=pr.restore_dual(reduced_dual)
            result={'status':r['status'],'iterations':r['iterations'],'solver_seconds':r['solver_seconds'],
                'backend_solve_seconds':r['backend_solve_seconds'],'audit':r,'original_dual':dual_metrics(pr,z),
                'dual_restore':restoration,'transform':{'route':args.route,'objective_scale':1.,'coordinate_scale':1.},
                'core_api_third_return_is_reduced_dual':True}
            extra={'reduced_dual':reduced_dual}
        else:result,x,z,extra=direct_transformed_solve(pr,args.route,args.seconds)
        assert program_digest(pr.original)==original_hash
        core.np.savez_compressed(out/'vectors.npz',x=x,original_dual=z,fs=1.,mode='upper',q=8,**extra)
        core.np.savez_compressed(out/'mesh.npz',xy=mesh.xy,tri=mesh.tri,material_id=mesh.material_id)
        record.update(status='COMPLETED',result=result,original_program_unchanged=True,
            output_hashes={p.name:sha(p.read_bytes())for p in sorted(out.glob('*.npz'))})
    except Exception as exc:
        record.update(status='ERROR',error=repr(exc),traceback=traceback.format_exc())
    finally:
        after=frozen_hashes()
        record.update(source_data_hashes_after=after,source_data_unchanged=before==after,
            finished_utc=utc(),process_elapsed_seconds=time.perf_counter()-started)
        write(out/'result.json',record)
    return 0 if record['status']=='COMPLETED' else 1

def maxdiff(a,b):return float(core.np.max(abs((a-b).data),initial=0.)) if core.sparse.issparse(a) else float(core.np.max(abs(a-b),initial=0.))

def parent(args):
    out=args.output.resolve()
    if out.exists():raise FileExistsError(out)
    out.mkdir(parents=True)
    before=frozen_hashes();started=time.perf_counter()
    summary={'status':'RUNNING','started_utc':utc(),'pid':os.getpid(),'fs':1.,'q':8,'optimizer_calls':0,
        'runs':[],'scope':'Six bounded fixed-point Clarabel upper-LP comparisons; no Fs root or GAP target claim.'}
    write(out/'summary.json',summary)
    routes=[('P02A','none'),('P02A','affine'),('D05A','none'),('D05A','affine'),('D05A','normalized_objective'),('D05A','velocity_units')]
    for case,route in routes:
        name=case+'_'+route;folder=out/name
        cmd=[sys.executable,'-X','utf8',str(Path(__file__).resolve()),'--child','--case',case,'--route',route,'--seconds',str(args.seconds),'--output',str(folder)]
        item={'name':name,'argv':cmd,'started_utc':utc()}
        clock=time.perf_counter()
        process=subprocess.Popen(cmd,stdout=subprocess.PIPE,stderr=subprocess.PIPE,text=True,encoding='utf-8')
        item['pid']=process.pid
        try:
            stdout,stderr=process.communicate(timeout=60)
            item['timeout']=False
        except subprocess.TimeoutExpired:
            process.kill();stdout,stderr=process.communicate();item['timeout']=True
        item.update(exit_code=process.returncode,finished_utc=utc(),process_wall_seconds=time.perf_counter()-clock,
            stdout=stdout,stderr=stderr,process_exited=True,result=str(folder.relative_to(ROOT)/'result.json'))
        if (folder/'result.json').exists():
            r=json.loads((folder/'result.json').read_text());item['result_status']=r['status'];summary['optimizer_calls']+=r['optimizer_calls']
        summary['runs'].append(item);write(out/'summary.json',summary)
        print(json.dumps({k:item.get(k)for k in ('name','exit_code','result_status','process_wall_seconds')}),flush=True)
    base,base_record=load_model('P02A');scaled,scaled_record=load_model('D05A')
    bp=core.Prepared(base,1.,'upper',8);sp=core.Prepared(scaled,1.,'upper',8)
    E,b,G,h,dims,c=bp.original;Es,bs,Gs,hs,dimss,cs=sp.original
    strain_rows=9*len(base.tri)
    law={'base_case':'P02A','scaled_case':'D05A','length_scale':1000.,'gamma_scale':.001,
        'tri_connectivity_equal':bool(core.np.array_equal(base.tri,scaled.tri)),
        'material_ids_equal':bool(core.np.array_equal(base.material_id,scaled.material_id)),
        'coordinates_x1000_max_abs_error':maxdiff(base.xy*1000,scaled.xy),
        'c_phi_psi_equal':bool(core.np.array_equal(base.materials[:,:3],scaled.materials[:,:3])),
        'gamma_div1000_equal':bool(core.np.array_equal(base.materials[:,3]/1000,scaled.materials[:,3])),
        'upper_original_E_max_abs_error':maxdiff(E,Es),
        'upper_original_b_div1000_max_abs_error':maxdiff(b/1000,bs),
        'upper_original_strain_G_div1000_max_abs_error':maxdiff(G[:strain_rows]/1000,Gs[:strain_rows]),
        'upper_original_jump_G_max_abs_error':maxdiff(G[strain_rows:],Gs[strain_rows:]),
        'upper_original_cost_x1000_max_abs_error':maxdiff(c*1000,cs),
        'q8_cone_layout_equal':dims==dimss,'primal_similarity_maps':[]}
    for route in ('none','affine'):
        p=out/('P02A_'+route)/'vectors.npz'
        if p.exists():
            with core.np.load(p,allow_pickle=False)as data:x=data['x'].copy();z=data['original_dual'].copy()
            mapped_x=x/1000
            mapped_z=z*1000;mapped_z[len(b):len(b)+strain_rows]*=1000
            metrics=core.evaluate(sp,mapped_x,mapped_z)
            core.np.savez_compressed(out/f'P02A_{route}_mapped_to_D05A.npz',x=mapped_x,original_dual=mapped_z,fs=1.,mode='upper',q=8)
            law['primal_similarity_maps'].append({'source_route':route,'original_q8_D05A_audit':metrics,
                'original_dual':dual_metrics(sp,mapped_z),'primal_scale':.001,
                'dual_equality_and_jump_scale':1000.,'dual_strain_scale':1000000.,
                'scope':'Exact positive similarity coordinate map; residuals and fields rechecked on original D05A, not accepted by a relative gate.'})
    summary.update(status='COMPLETED'if all(r.get('exit_code')==0 for r in summary['runs'])else'EXPERIMENT_HAS_ERRORS',
        finished_utc=utc(),process_elapsed_seconds=time.perf_counter()-started,
        source_data_hashes_before=before,source_data_hashes_after=frozen_hashes(),
        source_data_unchanged=before==frozen_hashes(),similarity_law=law,
        no_product_source_modified=True,no_native_or_VBA_execution=True,no_root_search=True,
        remaining_limits='This recorded 214-element geometry differs from the issue reporter current 400-element mesh. Cannot assign the DLL/helper failure from Clarabel diagnostics alone.')
    write(out/'summary.json',summary)
    return 0 if summary['status']=='COMPLETED'else 1

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--child',action='store_true');p.add_argument('--case',choices=['P02A','D05A']);p.add_argument('--route',choices=['none','affine','normalized_objective','velocity_units']);p.add_argument('--seconds',type=float,default=20.)
    p.add_argument('--output',type=Path,required=True);args=p.parse_args()
    if not 0<args.seconds<=30:raise ValueError('Bounded experiment requires 0 < seconds <= 30')
    raise SystemExit(child(args)if args.child else parent(args))
