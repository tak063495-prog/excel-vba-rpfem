"""Experimental prototype: full-vector phi=0 discrete root transport.

This is not called by the production accuracy driver. Adoption requires a
valid original source dual; a rejected source never produces a completed root.

Restricted to the existing dry, reference-only all-phi=psi=0 eligibility. Each
endpoint gets a fresh uncondensed program and the ordinary physical/raw dual
audit. Scalar metrics are never used as a substitute for the full dual vector.
"""
from __future__ import annotations
import argparse
import json
from pathlib import Path
import time

import accuracy_driver as driver
from experiment_core import np,Prepared,evaluate,write,LOWER_TARGET,UPPER_TARGET
from zero_phi_transport import require_reference_phi_zero


def sparse_difference(a,b):
    if a.shape!=b.shape:return float('inf')
    difference=a-b
    return float(np.max(np.abs(difference.data),initial=0.))


def verify_scaling(source,fresh):
    E,b,G,h,dims,cost=source.original;EE,bb,GG,hh,dd,cc=fresh.original
    ratio=source.fs/fresh.fs;mode=source.mode
    differences={'E':sparse_difference(E,EE),'G':sparse_difference(G,GG),
        'b':float(np.max(np.abs(b-bb),initial=0.)),'dims_match':dims==dd}
    if mode=='lower':
        differences['b_zero']=bool(np.all(b==0) and np.all(bb==0))
        differences['cost']=float(np.max(np.abs(cost-cc),initial=0.))
        differences['h_scaling']=float(np.max(np.abs(hh-h*ratio),initial=0.))
        exact=bool(differences['b_zero'] and differences['cost']==0.)
        scaled=differences['h_scaling']<=2e-13*(1+np.max(np.abs(hh),initial=0.))
    else:
        differences['h']=float(np.max(np.abs(h-hh),initial=0.))
        differences['cost_scaling']=float(np.max(np.abs(cc-cost*ratio),initial=0.))
        exact=differences['h']==0.
        scaled=differences['cost_scaling']<=2e-13*(1+np.max(np.abs(cc),initial=0.))
    if not (differences['E']==differences['G']==differences['b']==0. and differences['dims_match'] and exact and scaled):
        raise ValueError(f'Original program homogeneity check failed: {differences}')
    return differences


def signed_transport(mesh,fs,field,dual,mode,tolerance=1e-4):
    started=time.perf_counter();require_reference_phi_zero(mesh)
    if mode not in ['lower','upper'] or not np.isfinite(fs) or fs<=0:raise ValueError('Invalid mode or source Fs')
    if not 0<tolerance<=1e-4:raise ValueError('The final root tolerance cannot be relaxed')
    source=Prepared(mesh,fs,mode,8,reduction='none');E,b,G,h,dims,cost=source.original
    field=np.asarray(field).copy();dual=np.asarray(dual).copy()
    if field.ndim!=1 or len(field)!=len(cost) or dual.ndim!=1 or len(dual)!=len(b)+len(h) or not np.isfinite(dual).all():
        raise ValueError('Full finite original field and dual vectors are required')
    before=evaluate(source,field,dual)
    if not before['physical_pass'] or before['dual_stationarity']>1e-7 or before['dual_cone_error']>1e-8 or not driver.finite(before['dual_objective'],before['value']):
        raise ValueError(f'The source original physical or dual audit failed: {before}')
    if mode=='lower':
        expected=np.zeros(18*len(mesh.tri)+1);expected[-1]=-1.
        if not np.array_equal(cost,expected):raise ValueError('Free-load lower objective/layout required')
        lo=fs*before['value']/(LOWER_TARGET+2e-7)
        hi=fs*(-before['dual_objective'])/(LOWER_TARGET-3e-6)
    else:
        lo=fs*before['dual_objective']/(UPPER_TARGET+4e-6)
        hi=fs*before['value']/(UPPER_TARGET-2e-7)
    if not driver.finite(lo,hi) or not 0<lo<hi:raise ValueError('Transport did not produce an ordered positive bracket')
    endpoints=[]
    for endpoint in [lo,hi]:
        fresh=Prepared(mesh,endpoint,mode,8,reduction='none');scaling=verify_scaling(source,fresh)
        ratio=fs/endpoint
        xx=field*ratio if mode=='lower' else field.copy()
        zz=dual.copy() if mode=='lower' else dual*ratio
        audit=evaluate(fresh,xx,zz);sign=driver.sign(audit,mode)
        endpoints.append({'fs':float(endpoint),'field':xx,'dual':zz,'audit':audit,'sign':sign,'original_scaling':scaling})
    width=hi-lo;complete=bool([p['sign'] for p in endpoints]==[1,-1] and width<=tolerance*max(1.,(lo+hi)/2))
    selected=endpoints[0 if mode=='lower' else 1]
    complete=bool(complete and selected['audit']['physical_pass'] and selected['audit']['witness'])
    root={'mode':mode,'q':8,'case':mesh.provenance['case'],'policy':mesh.provenance['psi_policy'],
        'mesh_id':mesh.mesh_id,'physical_model_id':mesh.physical_model_id,'elements':len(mesh.tri),
        'root_complete':complete,'bracket':[float(lo),float(hi)],'bracket_width':float(width),
        'fs_tolerance':tolerance,'endpoint':selected['fs'],'bracket_signs':[p['sign'] for p in endpoints],
        'trials':2,'source_fs':fs,'seconds':time.perf_counter()-started,
        'method':'fresh_original_primal_dual_phi_zero_homogeneity',
        'opposite_numerical_endpoint_is_continuum_bound':False}
    selected['audit']['discrete_root_search']=root
    return {'source_audit':before,'root':root,'endpoints':endpoints,'selected':selected}


def run(case,mesh_path,fs,mode,output,seconds=120.):
    rec=json.loads((driver.CATALOG/f'{case}.json').read_text(encoding='utf8'))
    if rec['source_case'].get('pore'):raise ValueError('A coupled pore snapshot is required')
    mesh=driver.load_saved_mesh(mesh_path,rec);require_reference_phi_zero(mesh)
    output.mkdir(parents=True,exist_ok=True);driver.save_mesh(output/'original_mesh.npz',mesh)
    pr=Prepared(mesh,fs,mode,8,reduction='affine')
    result,field,_=driver.core.solve(pr,{'equilibrate_enable':False,'time_limit':seconds},max_iter=350,
        primal_polish=True,save=output/'source_full_vectors.npz')
    with np.load(output/'source_full_vectors.npz',allow_pickle=False) as data:dual=data['z'].copy()
    write(output/'source_solve.json',result)
    evidence=signed_transport(mesh,fs,field,dual,mode)
    for i,point in enumerate(evidence['endpoints']):
        np.savez_compressed(output/f'endpoint_{i}_full_vectors.npz',x=point['field'],z=point['dual'],fs=point['fs'],mode=mode,q=8)
        write(output/f'endpoint_{i}_audit.json',{'fs':point['fs'],'sign':point['sign'],'original_scaling':point['original_scaling'],'audit':point['audit']})
    selected=evidence['selected'];cert=driver.Certificate(mode,selected['fs'],mesh,selected['field'],
        selected['audit'],str(output/'source_full_vectors.npz'),evidence['root']['root_complete'])
    metadata=cert.publish(output)
    write(output/'import_audit.json',{'case':case,'certificates':{mode:metadata},'full_vector_transport_root':evidence['root']})
    write(output/'transport.json',{'source_audit':evidence['source_audit'],'root':evidence['root']})
    print('SIGNED_PHI_ZERO',json.dumps(evidence['root']),flush=True)
    return evidence


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--case',required=True);parser.add_argument('--mesh',type=Path,required=True)
    parser.add_argument('--fs',type=float,required=True);parser.add_argument('--mode',choices=['lower','upper'],default='lower')
    parser.add_argument('--output',type=Path,required=True);parser.add_argument('--seconds',type=float,default=120.)
    args=parser.parse_args();run(args.case,args.mesh,args.fs,args.mode,args.output,args.seconds)
