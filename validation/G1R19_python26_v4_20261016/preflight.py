"""Read-only numerical assembly/transport checks; starts no optimizer."""
from pathlib import Path
import argparse
import hashlib
import json
import sys
import time

ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'python'))
import accuracy_driver as driver
import experiment_core as core
import adaptive_policy as policy
from mesh_accuracy import lower_stress_score
np=core.np


def read_field(folder,mode,record):
    names=[f'best_{mode}_{suffix}' for suffix in ('mesh.npz','field.npz','certificate.json')]
    for retry in range(3):
        raw={name:(folder/name).read_bytes() for name in names}
        mesh=driver.load_saved_mesh(folder/names[0],record)
        with np.load(folder/names[1]) as data:
            field,fs=data['x'].copy(),float(data['fs'])
        metadata=json.loads(raw[names[2]].decode('utf-8'))
        if (all((folder/name).read_bytes()==value for name,value in raw.items())
                and metadata['mesh_id']==mesh.mesh_id and metadata['fs']==fs):
            return mesh,field,fs,metadata,{name:hashlib.sha256(value).hexdigest() for name,value in raw.items()}
    raise ValueError('Source changed while taking a read-only checkpoint snapshot')


def run(args):
    args.output.mkdir(parents=True,exist_ok=False)
    started=time.perf_counter();before=driver.source_hashes();rows={}
    for case in ('P06A','P06N'):
        _,rec=driver.load_model(case)
        folder=args.input/'cases'/case/'generic_fallback'/case
        mesh,field,fs,metadata,hashes=read_field(folder,'upper',rec)
        audit=core.evaluate(core.Prepared(mesh,fs,'upper',8),field)
        if not audit['physical_pass'] or not audit['witness']:raise ValueError('Fresh source upper rejected')
        raw=core.indicator(mesh,field)
        mixed,scoremeta=policy.upper_refinement_score(mesh,field,fs,.2)
        rawchild,rawchange=core.local_refine(mesh,rec,raw,.6)
        candidate,change=core.local_refine(mesh,rec,mixed,.8)
        old=core.prolong(mesh,candidate,field,'upper')
        new=policy.stable_prolong(mesh,candidate,field,'upper')
        program=core.Prepared(candidate,fs,'upper',8)
        oldaudit=core.evaluate(program,old);newaudit=core.evaluate(program,new)
        driver.save_mesh(args.output/(case+'_upper_candidate_mesh.npz'),candidate)
        np.savez_compressed(args.output/(case+'_upper_candidate_field.npz'),x=new,fs=fs,mode='upper',q=8)
        npath=args.native_mesh/case/'target_1000_mesh.npz'
        native=driver.load_saved_mesh(npath,rec)
        lower,lfield,lfs,lmetadata,lhashes=read_field(folder,'lower',rec)
        laudit=core.evaluate(core.Prepared(lower,lfs,'lower',8),lfield)
        if not laudit['physical_pass'] or not laudit['witness']:raise ValueError('Fresh source lower rejected')
        lscore=lower_stress_score(lower,lfield,lfs)
        lchild,lchange=core.local_refine(lower,rec,lscore,.8)
        lnew=policy.stable_prolong(lower,lchild,lfield,'lower')
        lnewaudit=core.evaluate(core.Prepared(lchild,lfs,'lower',8),lnew)
        rows[case]=dict(source_upper=dict(fs=fs,elements=len(mesh.tri),root_complete=metadata['root_complete'],audit=audit,sha256=hashes),
            upper_score=scoremeta,original_kinematic_refinement=rawchange,mixed_refinement=change,
            old_global_coordinate_transfer=oldaudit,new_local_coordinate_transfer=newaudit,
            maximum_transfer_difference=float(np.max(abs(old-new))),
            source_lower=dict(fs=lfs,elements=len(lower.tri),root_complete=lmetadata['root_complete'],audit=laudit,sha256=lhashes),
            lower_refinement=lchange,lower_transported_audit=lnewaudit,
            fresh_native_mesh=dict(path=str(npath),sha256=hashlib.sha256(npath.read_bytes()).hexdigest(),elements=len(native.tri),
                mesh_id=native.mesh_id,physical_model_id=native.physical_model_id,
                canonical_model_identity_exact=native.physical_model_id==mesh.physical_model_id,
                areas_by_material=np.bincount(native.material_id,weights=core.topology(native)[0]).tolist()),
            publishes_bound=False,publishes_root=False)
    after=driver.source_hashes()
    if before!=after:raise ValueError('Implementation changed during preflight')
    report=dict(status='PASS',scope='fresh-field read-only diagnostic; no optimizer or root search, no new bound',
        cases=rows,seconds=time.perf_counter()-started,source_sha256_before=before,source_sha256_after=after)
    driver.write(args.output/'PREFLIGHT.json',report)
    for case,row in rows.items():
        print(case,'upper',row['source_upper']['elements'],'->',row['mixed_refinement']['elements_after'],
              'transport physical',row['old_global_coordinate_transfer']['physical_pass'],row['new_local_coordinate_transfer']['physical_pass'],
              'lower',row['source_lower']['elements'],'->',row['lower_refinement']['elements_after'],
              row['lower_transported_audit']['physical_pass'],flush=True)
    print('NO_SOLVER_PREFLIGHT',report['status'],report['seconds'],flush=True)


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--input',type=Path,required=True)
    p.add_argument('--native-mesh',type=Path,required=True);p.add_argument('--output',type=Path,required=True)
    run(p.parse_args())
