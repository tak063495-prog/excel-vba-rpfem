"""Reproduce direct original-dual experiments from an explicit saved mesh."""
import argparse
import hashlib
import io
import json
from pathlib import Path

import accuracy_driver as driver
from experiment_core import np,Prepared,write
from original_dual import solve_original_dual


def run(case,mesh_path,fs,mode,output,seconds,compare=False):
    catalog=driver.CATALOG/f'{case}.json';catalog_bytes=catalog.read_bytes();rec=json.loads(catalog_bytes.decode('utf8'))
    if rec['source_case'].get('pore'):raise ValueError('A coupled pore snapshot is required')
    raw=mesh_path.read_bytes()
    with np.load(io.BytesIO(raw),allow_pickle=False) as data:
        mesh=driver.core.mesh_from_arrays(data['xy'].copy(),data['tri'].copy(),data['material_id'].copy(),rec)
    output.mkdir(parents=True,exist_ok=True);driver.save_mesh(output/f'original_{mode}_mesh.npz',mesh)
    pr=Prepared(mesh,fs,mode,8,reduction='none');rows=[]
    configurations=[(True,0.,False)]
    if compare:configurations=[(False,0.,False),(True,0.,False),(False,0.,True),(True,0.,True),(False,1e-10,False),(True,1e-10,False)]
    for eq,reg,scale in configurations:
        evidence=solve_original_dual(pr,seconds,return_dual=True,equilibrate=eq,regularization=reg,equality_scale=scale)
        dual=evidence.pop('dual',np.zeros(0));stem=f'candidate_eq{int(eq)}_reg{reg:g}_scale{int(scale)}'
        np.savez_compressed(output/(stem+'_dual.npz'),dual=dual,fs=fs,mode=mode,q=8)
        write(output/(stem+'_audit.json'),evidence);rows.append(evidence)
        write(output/'diagnosis.json',{'case':case,'mode':mode,'fs':fs,'q':8,'source_mesh':str(mesh_path),
            'source_mesh_sha256':hashlib.sha256(raw).hexdigest(),'catalog_sha256':hashlib.sha256(catalog_bytes).hexdigest(),
            'helper_sha256':hashlib.sha256(Path(__file__).with_name('original_dual.py').read_bytes()).hexdigest(),
            'trials':rows,'primal_bound_published':False})
        print('ORIGINAL_DUAL',json.dumps(evidence),flush=True)
    return rows


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--case',required=True)
    parser.add_argument('--mesh',type=Path,required=True);parser.add_argument('--fs',type=float,required=True)
    parser.add_argument('--mode',choices=['lower','upper'],default='upper');parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--seconds',type=float,default=90.);parser.add_argument('--compare',action='store_true')
    args=parser.parse_args();run(args.case,args.mesh,args.fs,args.mode,args.output,args.seconds,args.compare)
