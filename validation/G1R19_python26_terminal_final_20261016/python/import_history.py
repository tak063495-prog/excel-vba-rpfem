"""Import and re-audit the two best same-model bounds from the user's archive."""
import argparse
import json
from pathlib import Path
import accuracy_driver as driver
from experiment_core import np, Prepared, evaluate, write


def run(history, output):
    followup = history / '02_Followup_Experiments'
    adoption = json.loads((followup / 'results/adoption_summary.json').read_text(encoding='utf8'))
    for record in adoption['cases']:
        case = record['case']
        _, rec = driver.load_model(case)
        folder = output / case
        folder.mkdir(parents=True, exist_ok=True)
        certificates = {}
        for mode in ['lower', 'upper']:
            source = record[mode + '_source']
            stem = source['label'].removesuffix('_' + mode)
            mesh = driver.load_saved_mesh(followup / 'results' / (stem + '_mesh.npz'), rec)
            with np.load(followup / source['field_path']) as data:
                field, fs = data['x'].copy(), float(data['fs'])
            result = evaluate(Prepared(mesh, fs, mode, 8), field)
            assert result['physical_pass'] and result['witness'], (case, mode, result)
            old=json.loads((followup/'results'/(source['label']+'_summary.json')).read_text(encoding='utf8'))
            trials=json.loads((followup/'results'/(source['label']+'_trials.json')).read_text(encoding='utf8'))
            endpoints=[next((r for r in trials if r.get('fs')==old.get(key)),{}) for key in ['root_search_lo','root_search_hi']]
            signs=[driver.sign(r,mode) if r else None for r in endpoints]
            complete=bool(old.get('complete') and old.get('fs_endpoint')==fs and
                          old.get('root_search_width',1)<=1e-4*max(1.,fs) and signs==[1,-1] and
                          all(r.get('q')==8 and r.get('mode')==mode and r.get('case')==case for r in endpoints))
            result['discrete_root_search']={'root_complete':complete,'endpoint':fs,
                      'bracket':[old.get('root_search_lo'),old.get('root_search_hi')],
                      'bracket_width':old.get('root_search_width'),'fs_tolerance':1e-4,
                      'bracket_signs':signs,'q':8,'mode':mode,'case':case,'policy':mesh.provenance['psi_policy'],
                      'mesh_id':mesh.mesh_id,'physical_model_id':mesh.physical_model_id,
                      'source':source['label']+'_summary.json'}
            certificate = driver.Certificate(mode, fs, mesh, field, result, source['label'],complete)
            certificates[mode] = certificate.publish(folder)
        write(folder / 'import_audit.json', {'case': case, 'certificates': certificates,
              'scope': 'Independent Windows re-audit of archived fields; not a new root search'})
        print('IMPORTED', case, {side: cert['fs'] for side, cert in certificates.items()}, flush=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser()
    parser.add_argument('--history', required=True, type=Path)
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args()
    run(args.history.resolve(), args.output.resolve())
