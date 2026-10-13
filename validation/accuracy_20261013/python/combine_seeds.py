"""Select independently re-audited endpoints across same-model experiments."""
import argparse
import json
from pathlib import Path
import accuracy_driver as driver
from experiment_core import np, Prepared, evaluate, write


def run(case, sources, output):
    _,rec=driver.load_model(case)
    best={}
    candidates=[]
    for source in sources:
        folder=source/case if (source/case).is_dir() else source
        if (folder/'summary.json').exists():
            metadata=json.loads((folder/'summary.json').read_text(encoding='utf8'))
        elif (folder/'import_audit.json').exists():
            metadata=json.loads((folder/'import_audit.json').read_text(encoding='utf8'))['certificates']
        else:
            continue
        for mode in ['lower','upper']:
            if mode not in metadata:
                continue
            old=metadata[mode]
            mesh=driver.load_saved_mesh(folder/f'best_{mode}_mesh.npz',rec)
            with np.load(folder/f'best_{mode}_field.npz') as data:
                if str(data['mode'])!=mode or int(data['q'])!=8:
                    raise ValueError('Stored mode or quadrature differs')
                field,fs=data['x'].copy(),float(data['fs'])
            if fs!=old['fs'] or mesh.mesh_id!=old['mesh_id'] or mesh.physical_model_id!=old['physical_model_id']:
                raise ValueError('Stored field and certificate differ')
            result=evaluate(Prepared(mesh,fs,mode,8),field)
            candidates.append({'source':str(folder),'mode':mode,'fs':fs,'witness':result['witness']})
            if not result['witness']:
                continue
            root=old.get('audit',{}).get('discrete_root_search')
            if root:
                result['discrete_root_search']=root
            cert=driver.Certificate(mode,fs,mesh,field,result,str(folder),bool(old.get('root_complete',False)))
            if mode not in best or (fs>best[mode].fs if mode=='lower' else fs<best[mode].fs):
                best[mode]=cert
    if set(best)!={'lower','upper'} or not 0<best['lower'].fs<=best['upper'].fs:
        raise ValueError('No ordered pair of physical witnesses')
    folder=output/case
    folder.mkdir(parents=True,exist_ok=True)
    certs={mode:cert.publish(folder) for mode,cert in best.items()}
    width=2*(best['upper'].fs-best['lower'].fs)/(best['lower'].fs+best['upper'].fs)
    write(folder/'import_audit.json',{'case':case,'certificates':certs,'candidates':candidates,
         'relative_width_symmetric':width,
         'scope':'Original physical fields re-audited; driver must validate retained root metadata'})
    print('COMBINED',case,best['lower'].fs,best['upper'].fs,width,flush=True)


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--cases',nargs='+',required=True)
    parser.add_argument('--sources',nargs='+',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args()
    for case in args.cases:
        run(case,[p.resolve() for p in args.sources],args.output.resolve())
