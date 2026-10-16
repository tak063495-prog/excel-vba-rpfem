"""Reassemble original q8 programs in a fresh process; never run a solver."""
from pathlib import Path
import argparse
import hashlib
import json
import sys
from decimal import Decimal,ROUND_HALF_UP

ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'python'))

def audit(case, folder, target=.02):
    import benchmark_models
    benchmark_models.install_model_builder()
    import accuracy_driver as driver
    import report_accuracy as report
    _,record=benchmark_models.load_model(case)
    result={'case':case,'scope':'Fresh original q8 primal audit; retained own-root identity and signs checked; solver_calls=0',
            'target_gap':target,'root_tolerance':1e-4,'sides':{},'errors':[],'physical_gap_pass':False,'root_complete':False,'target_met':False}
    accepted={}
    for mode in ('lower','upper'):
        path=folder/f'best_{mode}_certificate.json'
        try:
            candidate=report.collect_candidate(path,mode,case,record,[folder],1e-4)
            if candidate is None:raise ValueError('MISSING_CERTIFICATE')
            accepted[mode]=candidate
            result['sides'][mode]={
                'fs':candidate['fs'],'q':8,'elements':len(candidate['mesh'].tri),'mesh_id':candidate['mesh'].mesh_id,
                'physical_model_id':candidate['mesh'].physical_model_id,'policy':candidate['mesh'].provenance['psi_policy'],
                'physical_pass':True,'witness':True,'root_complete':candidate['root_complete'],'audit':candidate['audit'],
                'file_sha256':candidate['source_hashes']}
        except (OSError,ValueError,KeyError,AssertionError,TypeError,IndexError,RuntimeError) as error:
            result['errors'].append({'mode':mode,'error':f'{type(error).__name__}: {error}'})
    if len(accepted)==2:
        lower,upper=accepted['lower'],accepted['upper']
        if lower['mesh'].physical_model_id!=upper['mesh'].physical_model_id or not 0<lower['fs']<=upper['fs']:
            result['errors'].append({'error':'INCONSISTENT_MODEL_OR_REVERSED_BOUNDS'})
        else:
            width=2*(upper['fs']-lower['fs'])/(upper['fs']+lower['fs'])
            result.update(lower_fs=lower['fs'],upper_fs=upper['fs'],relative_gap=width,
                          physical_gap_pass=width<=target,root_complete=lower['root_complete'] and upper['root_complete'])
            lower_round=str(Decimal(str(lower['fs'])).quantize(Decimal('.01'),rounding=ROUND_HALF_UP))
            upper_round=str(Decimal(str(upper['fs'])).quantize(Decimal('.01'),rounding=ROUND_HALF_UP))
            result.update(lower_rounded_2dp=lower_round,upper_rounded_2dp=upper_round,rounded_2dp_agree=lower_round==upper_round)
            result['target_met']=result['physical_gap_pass'] and result['root_complete']
    result['status']='VERIFIED_2PCT' if result['target_met'] else 'PHYSICAL_GAP_ONLY' if result['physical_gap_pass'] else 'INCOMPLETE' if len(accepted)<2 else 'GAP_NOT_MET'
    return result

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--case',required=True);ap.add_argument('--folder',type=Path,required=True)
    ap.add_argument('--target-gap',type=float,default=.02);ap.add_argument('--output',type=Path,required=True);args=ap.parse_args()
    if not 0<args.target_gap<2:raise ValueError('Invalid target gap')
    result=audit(args.case,args.folder.resolve(),args.target_gap)
    args.output.write_text(json.dumps(result,indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8')
    print(json.dumps({k:result.get(k) for k in ('case','status','relative_gap','root_complete','target_met')},ensure_ascii=False),flush=True)

if __name__=='__main__':main()
