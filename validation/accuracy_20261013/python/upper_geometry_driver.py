"""Audit a flow-metric flip candidate on its own upper mesh and root.

Run with --case P06A --seed PATH --output PATH. The source witness retains its
own mesh. Interpolation is only a warm candidate after the full physical gate.
"""
from __future__ import annotations
import argparse
import hashlib
import json
from pathlib import Path
import platform
import accuracy_driver as driver
import experiment_core as core
from mesh_geometry import flow_metric_flip, interpolate_velocity

SOURCE_HASHES=dict(driver.SOURCE_HASHES)
for source in [Path(__file__),Path(__file__).with_name('mesh_geometry.py')]:
    SOURCE_HASHES[str(source.resolve().relative_to(driver.REPO))]=hashlib.sha256(source.read_bytes()).hexdigest()


def publish(certificate, folder, case, source_fs, root=None, change=None):
    upper=certificate.publish(folder)
    core.write(folder/'import_audit.json',{'case':case,'certificates':{'upper':upper}})
    core.write(folder/'summary.json',{'case':case,'upper':upper,'upper_only':True,
               'one_percent_target_not_evaluated':True,'source_fs_upper':source_fs,
               'stronger_than_source':certificate.fs<source_fs,'root':root,'mesh_change':change,
               'environment':{'python':platform.python_version(),'numpy':core.np.__version__,
                              'clarabel':core.clarabel.__version__,'threads':1,'source_sha256':SOURCE_HASHES}})


def run(args):
    if args.passes<0 or not driver.finite(args.tolerance,args.solve_seconds) or not 0<args.tolerance<=1e-4 or args.solve_seconds<=0:
        raise ValueError('Pass count, root tolerance, or solve budget is invalid')
    _,record=driver.load_model(args.case)
    parent=driver.load_saved_mesh(args.seed/'best_upper_mesh.npz',record)
    with core.np.load(args.seed/'best_upper_field.npz',allow_pickle=False) as data:
        if str(data['mode'])!='upper' or int(data['q'])!=8:raise ValueError('Mode or quadrature mismatch')
        velocity,fs=data['x'].copy(),float(data['fs'])
    if not driver.finite(fs) or fs<=0:raise ValueError('Invalid source strength factor')
    source=core.evaluate(core.Prepared(parent,fs,'upper',8,reduction='none'),velocity)
    if not source['physical_pass'] or not source['witness']:raise ValueError('Source failed the full original physical gate')
    certificate_path=args.seed/'best_upper_certificate.json'
    if certificate_path.is_file():metadata=json.loads(certificate_path.read_text(encoding='utf8'))
    else:
        path=args.seed/('summary.json' if (args.seed/'summary.json').is_file() else 'import_audit.json')
        if not path.is_file():raise ValueError('Source certificate metadata is required')
        data=json.loads(path.read_text(encoding='utf8'));metadata=data.get('certificates',data).get('upper',{})
    if metadata.get('fs')!=fs or metadata.get('mesh_id')!=parent.mesh_id or metadata.get('physical_model_id')!=parent.physical_model_id or metadata.get('case')!=args.case or metadata.get('mode')!='upper' or metadata.get('q')!=8:
        raise ValueError('Source field, certificate, mesh, case, or model identity differs')
    source.update(case=args.case,mode='upper',q=8,fs=fs,mesh_id=parent.mesh_id,
                  physical_model_id=parent.physical_model_id,search_sign=-1)
    folder=args.output/args.case;folder.mkdir(parents=True,exist_ok=True)
    retained=driver.Certificate('upper',fs,parent,velocity,source,str(args.seed),False)
    parent_folder=folder/'source_witness';parent_folder.mkdir(parents=True,exist_ok=True)
    publish(retained,parent_folder,args.case,fs)
    child,change=flow_metric_flip(parent,record,velocity,args.passes)
    driver.save_mesh(folder/'candidate_mesh.npz',child)
    core.write(folder/'mesh_change.json',change)
    transferred=interpolate_velocity(parent,child,velocity,8)
    prepared=core.Prepared(child,fs,'upper',8,reduction='none')
    E,b,*_=prepared.original
    work_rows=core.np.flatnonzero(b!=0)
    if len(work_rows)!=1:raise ValueError('Upper reference-work normalization has an unexpected layout')
    row=int(work_rows[0]);work=float((E[row]@transferred)[0]/b[row])
    if driver.finite(work) and work>0:transferred/=work
    audit=core.evaluate(prepared,transferred)
    audit.update(case=args.case,mode='upper',q=8,fs=fs,mesh_id=child.mesh_id,
                 physical_model_id=child.physical_model_id,work_normalization=work,
                 origin='nonnested_same_material_P2_interpolation')
    core.write(folder/'interpolated_audit.json',audit)
    warm=None
    if audit['physical_pass'] and audit['witness']:
        audit['search_sign']=-1
        warm=driver.Certificate('upper',fs,child,transferred,audit,'fully audited nonnested interpolation',False)
        publish(warm,folder,args.case,fs,change=change)
    else:
        core.write(folder/'summary.json',{'case':args.case,'upper_only':True,'one_percent_target_not_evaluated':True,
                   'reason':'INTERPOLATED_FIELD_IS_NOT_AN_ADMISSIBLE_UPPER_BOUND','mesh_change':change})
    delta=max(.02,.04*fs)
    bracket=args.bracket or (max(1e-5,fs-delta),fs+delta)
    print('FLOW_MESH',args.case,len(child.tri),change['flips'],'warm',warm is not None,flush=True)
    candidate,root=driver.search(child,'upper',bracket,folder,args.case+'_flow_metric_upper',
                                 args.tolerance,args.solve_seconds,warm)
    if candidate is not None:
        publish(candidate,folder,args.case,fs,root,change)
        print('FLOW_UPPER',args.case,candidate.fs,candidate.root_complete,'improved',candidate.fs<fs,flush=True)
    else:
        core.write(folder/'summary.json',{'case':args.case,'upper_only':True,'root':root,
                   'reason':'NO_ORIGINAL_PHYSICAL_UPPER_WITNESS','mesh_change':change})


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--case',required=True)
    parser.add_argument('--seed',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--passes',type=int,default=5)
    parser.add_argument('--bracket',type=float,nargs=2)
    parser.add_argument('--tolerance',type=float,default=1e-4)
    parser.add_argument('--solve-seconds',type=float,default=120)
    run(parser.parse_args())
