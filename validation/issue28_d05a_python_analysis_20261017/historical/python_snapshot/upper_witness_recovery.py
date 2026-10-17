"""Recheck an upper field on its own mesh, then complete its discrete root.

This CLI does not infer admissibility from interpolation or a solver status.
The imported witness must pass every original q=8 physical gate. An incomplete
root is retained explicitly until driver.search supplies both signed endpoints.
"""
from pathlib import Path
import argparse
import hashlib
import json

import accuracy_driver as driver
import experiment_core as core


def run(args):
    _,record=driver.load_model(args.case)
    if args.output.exists():
        raise ValueError('Use a new output directory to preserve prior evidence')
    mesh=driver.load_saved_mesh(args.mesh,record)
    with core.np.load(args.field,allow_pickle=False) as data:
        if str(data['mode'])!='upper' or int(data['q'])!=8:
            raise ValueError('Only q=8 upper fields can be recovered')
        x=data['x'].copy();fs=float(data['fs'])
    if not driver.finite(fs) or fs<=0 or not core.np.isfinite(x).all():
        raise ValueError('Source field and factor must be finite')
    result=core.evaluate(core.Prepared(mesh,fs,'upper',8,reduction='none'),x)
    if not result['physical_pass'] or not result['witness']:
        raise ValueError('Imported field fails an original physical upper gate')
    lo,hi=args.bracket
    if not driver.finite(lo,hi,args.solve_seconds,args.tolerance) or not 0<lo<fs<=hi:
        raise ValueError('Bracket must contain the audited source factor')
    if not 0<args.tolerance<=1e-4 or args.solve_seconds<=0:
        raise ValueError('Root tolerance or time budget is invalid')
    provenance={'method':'independent_original_upper_witness_recovery',
                'source_mesh':str(args.mesh.resolve()),'source_field':str(args.field.resolve()),
                'mesh_sha256':hashlib.sha256(args.mesh.read_bytes()).hexdigest(),
                'field_sha256':hashlib.sha256(args.field.read_bytes()).hexdigest(),
                'source_sha256':dict(driver.SOURCE_HASHES),
                'source_root_complete':False}
    provenance['source_sha256'][str(Path(__file__).resolve().relative_to(driver.REPO))]=hashlib.sha256(Path(__file__).read_bytes()).hexdigest()
    result.update(case=args.case,mode='upper',q=8,fs=fs,mesh_id=mesh.mesh_id,
                  physical_model_id=mesh.physical_model_id,search_sign=-1,
                  origin=provenance['method'])
    source=driver.Certificate('upper',fs,mesh,x,result,provenance['method'],False)
    folder=args.output/args.case;folder.mkdir(parents=True)
    source_folder=folder/'source_witness';source_folder.mkdir()
    imported=source.publish(source_folder)
    core.write(source_folder/'import_audit.json',{'case':args.case,'certificates':{'upper':imported},'provenance':provenance})
    best,root=driver.search(mesh,'upper',(lo,hi),folder,args.case+'_upper_recovery',
                            args.tolerance,args.solve_seconds,source)
    metadata=best.publish(folder) if best is not None else None
    core.write(folder/'summary.json',{'case':args.case,'upper':metadata,'upper_only':True,
               'one_percent_target_not_evaluated':True,'root':root,'provenance':provenance})
    if metadata is not None:
        core.write(folder/'import_audit.json',{'case':args.case,'certificates':{'upper':metadata},'provenance':provenance})
    print('RECOVERED_UPPER',args.case,None if best is None else best.fs,
          False if best is None else best.root_complete,flush=True)


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--case',required=True);parser.add_argument('--mesh',type=Path,required=True)
    parser.add_argument('--field',type=Path,required=True);parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--bracket',type=float,nargs=2,required=True)
    parser.add_argument('--solve-seconds',type=float,default=120.)
    parser.add_argument('--tolerance',type=float,default=1e-4)
    run(parser.parse_args())
