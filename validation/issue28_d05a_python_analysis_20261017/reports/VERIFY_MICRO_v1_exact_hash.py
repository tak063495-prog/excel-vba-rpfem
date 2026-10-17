"""Re-audit the new fixed-point vectors with zero optimization calls."""
from __future__ import annotations
import argparse
import json
from pathlib import Path
import sys
sys.dont_write_bytecode=True
from run_fixed_point_experiment import ROOT,core,dual_metrics,frozen_hashes,load_model,program_digest,sha,write

def main():
    p=argparse.ArgumentParser();p.add_argument('--run',type=Path,default=ROOT/'results/fixed_point_02');p.add_argument('--output',type=Path,required=True);args=p.parse_args()
    if args.output.exists():raise FileExistsError(args.output)
    root=args.run.resolve();before=frozen_hashes();records=[]
    summary=json.loads((root/'summary.json').read_text(encoding='utf-8'))
    assert summary['status']=='COMPLETED'and summary['optimizer_calls']==6
    for entry in summary['runs']:
        assert entry['exit_code']==0 and entry['process_exited']and not entry['timeout']
        folder=root/entry['name'];meta=json.loads((folder/'result.json').read_text(encoding='utf-8'))
        assert meta['status']=='COMPLETED'and meta['source_data_unchanged']
        assert meta['source_data_hashes_before']==before==meta['source_data_hashes_after']
        assert meta['fs']==1 and meta['q']==8
        for name,expected in meta['output_hashes'].items():assert sha((folder/name).read_bytes())==expected
        mesh,_=load_model(meta['case']);pr=core.Prepared(mesh,1.,'upper',8,reduction='none')
        assert program_digest(pr.original)==meta['original_program']['sha256']
        with core.np.load(folder/'vectors.npz',allow_pickle=False)as data:
            assert float(data['fs'])==1 and int(data['q'])==8 and str(data['mode'])=='upper'
            x,z=data['x'].copy(),data['original_dual'].copy()
        observed=core.evaluate(pr,x,z);expected=meta['result']['audit']
        for key in ('physical_pass','witness','numerical_gate'):assert observed[key]is expected[key],(entry['name'],key)
        for key in ('objective','value','eq_error','cone_error','dual_objective','dual_stationarity','dual_cone_error','dual_gap'):
            assert abs(observed[key]-expected[key])<=1e-12*max(1.,abs(expected[key])),(entry['name'],key)
        dual=dual_metrics(pr,z)
        assert dual==meta['result']['original_dual'],entry['name']
        records.append({'name':entry['name'],'status':'REPLAY_MATCH','original_program_sha256':program_digest(pr.original),
            'physical_pass':observed['physical_pass'],'original_dual_guide_accepted':dual['original_upper_inverse_guide_accepted'],
            'full_original_audit':observed,'full_original_dual':dual,'vectors_sha256':sha((folder/'vectors.npz').read_bytes())})
    output={'status':'PASS','optimizer_calls':0,'records':records,'source_data_unchanged':before==frozen_hashes(),
        'scope':'Reassembled original q8 program and independently implemented physical audit; compare to frozen six-run diagnostics. Rejects are deliberately retained. No new root or physical target certificate.',
        'auditor_sha256':sha(Path(__file__).read_bytes()),'producer_summary_sha256':sha((root/'summary.json').read_bytes())}
    write(args.output,output);print(json.dumps({'status':'PASS','records':len(records),'optimizer_calls':0,'output':str(args.output)}))
if __name__=='__main__':main()
