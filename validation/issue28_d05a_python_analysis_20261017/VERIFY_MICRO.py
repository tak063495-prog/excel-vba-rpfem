"""Re-audit the new fixed-point vectors with zero optimization calls."""
from __future__ import annotations
import argparse
import copy
import json
from pathlib import Path
import sys
sys.dont_write_bytecode=True
from run_fixed_point_experiment import ROOT,core,dual_metrics,frozen_hashes,load_model,program_digest,sha,write
from portable_program_audit import compare_program,compare_metric,unpack_program

EXPECTED_RUN_NAMES={'P02A_none','P02A_affine','D05A_none','D05A_affine','D05A_normalized_objective','D05A_velocity_units'}

def validate_run_summary(summary):
    assert summary['status']=='COMPLETED'and summary['optimizer_calls']==6
    assert summary['fs']==1.0 and summary['q']==8
    names=[entry['name']for entry in summary['runs']]
    assert len(names)==6 and len(set(names))==6 and set(names)==EXPECTED_RUN_NAMES,'SIX_UNIQUE_EXPECTED_ROUTES_REQUIRED'

def validate_run_entry(entry,meta):
    assert entry['name']in EXPECTED_RUN_NAMES
    assert entry['name']==meta['case']+'_'+meta['route'],'ENTRY_CASE_ROUTE_MISMATCH'

def bind_recorded_program_digest(recorded_digest,observed_program,reference_program):
    # Frozen runs retain their byte-bound references. A newly generated run
    # may instead match the current host exactly. An unknown third digest has
    # no coefficient evidence and must fail explicitly.
    if recorded_digest==program_digest(reference_program):return 'exact_frozen_reference_digest'
    if recorded_digest==program_digest(observed_program):return 'exact_current_assembly_digest'
    raise AssertionError('RECORDED_PROGRAM_DIGEST_HAS_NO_EXACT_REFERENCE_OR_CURRENT_ASSEMBLY')

def load_references():
    path=ROOT/'program_refs/PROGRAM_REFERENCE_MANIFEST.json'
    manifest=json.loads(path.read_text(encoding='utf-8'))
    assert manifest['schema']=='issue28_micro_original_program_references_v1'
    assert manifest['source_data_manifest_sha256']==sha((ROOT/'SOURCE_DATA_MANIFEST.json').read_bytes())
    assert manifest['producer_script_sha256']==sha((ROOT/'run_fixed_point_experiment.py').read_bytes())
    assert manifest['comparator_sha256']==sha((ROOT/'portable_program_audit.py').read_bytes())
    result={}
    for case,record in manifest['programs'].items():
        assert case in ('P02A','D05A')
        target=ROOT/record['path']
        assert target.stat().st_size==record['bytes']and sha(target.read_bytes())==record['sha256']
        with core.np.load(target,allow_pickle=False)as data:
            assert set(data.files)==set(record['array_layout'])
            arrays={name:data[name].copy()for name in data.files}
        for name,array in arrays.items():
            spec=record['array_layout'][name]
            assert list(array.shape)==spec['shape']and str(array.dtype)==spec['dtype']
        program=unpack_program(arrays)
        assert program_digest(program)==record['original_program_sha256']
        for old in record['frozen_producer_records']:
            meta=ROOT/old['path'];assert sha(meta.read_bytes())==old['sha256']
            assert old['recorded_original_program_sha256']==record['original_program_sha256']
        result[case]=(program,record)
    assert set(result)=={'P02A','D05A'}
    return result,sha(path.read_bytes())

def main():
    p=argparse.ArgumentParser();p.add_argument('--run',type=Path,default=ROOT/'results/fixed_point_02');p.add_argument('--output',type=Path,required=True);args=p.parse_args()
    if args.output.exists():raise FileExistsError(args.output)
    root=args.run.resolve();before=frozen_hashes();records=[]
    references,reference_manifest_sha256=load_references()
    summary=json.loads((root/'summary.json').read_text(encoding='utf-8'))
    validate_run_summary(summary)
    for entry in summary['runs']:
        assert entry['exit_code']==0 and entry['process_exited']and not entry['timeout']
        folder=root/entry['name'];meta=json.loads((folder/'result.json').read_text(encoding='utf-8'))
        validate_run_entry(entry,meta)
        assert meta['status']=='COMPLETED'and meta['source_data_unchanged']
        assert meta['source_data_hashes_before']==before==meta['source_data_hashes_after']
        assert meta['fs']==1 and meta['q']==8
        for name,expected in meta['output_hashes'].items():assert sha((folder/name).read_bytes())==expected
        mesh,_=load_model(meta['case']);pr=core.Prepared(mesh,1.,'upper',8,reduction='none')
        reference,reference_record=references[meta['case']]
        for key,value in (('nodes',len(mesh.xy)),('elements',len(mesh.tri)),('mesh_id',mesh.mesh_id),
                ('physical_model_id',mesh.physical_model_id),('policy',mesh.provenance['psi_policy'])):
            assert meta['model'][key]==value,(entry['name'],key)
        recorded_digest_binding=bind_recorded_program_digest(meta['original_program']['sha256'],pr.original,reference)
        coefficient_comparison=compare_program(pr.original,reference)
        with core.np.load(folder/'vectors.npz',allow_pickle=False)as data:
            assert float(data['fs'])==1 and int(data['q'])==8 and str(data['mode'])=='upper'
            x,z=data['x'].copy(),data['original_dual'].copy()
        # The old byte-bound program remains available for reference replay.
        # Rebuilt coefficients can differ in their last bits across OpenBLAS
        # dispatches; both programs must retain every original gate decision.
        reference_prepared=copy.copy(pr);reference_prepared.original=reference
        reference_observed=core.evaluate(reference_prepared,x,z)
        observed=core.evaluate(pr,x,z);expected=meta['result']['audit'];metric_comparisons=[]
        for source,actual in (('reference',reference_observed),('rebuilt',observed)):
            for key in ('physical_pass','witness','numerical_gate'):assert actual[key]is expected[key],(entry['name'],source,key)
            for key in ('objective','value','eq_error','cone_error','dual_objective','dual_stationarity','dual_cone_error','dual_gap'):
                metric_comparisons.append(compare_metric(actual[key],expected[key],source+'.'+key))
        dual=dual_metrics(pr,z)
        old_dual=meta['result']['original_dual']
        for key in ('original_upper_inverse_guide_accepted','thresholds','scope'):assert dual[key]==old_dual[key],(entry['name'],key)
        for key in ('stationarity','cone_error','objective'):metric_comparisons.append(compare_metric(dual[key],old_dual[key],'original_dual.'+key))
        records.append({'name':entry['name'],'status':'PORTABLE_REPLAY_MATCH','original_program_sha256':program_digest(pr.original),
            'recorded_original_program_sha256':meta['original_program']['sha256'],
            'recorded_program_digest_binding':recorded_digest_binding,
            'reference_npz_sha256':reference_record['sha256'],'coefficient_comparison':coefficient_comparison,
            'metric_comparisons':metric_comparisons,
            'physical_pass':observed['physical_pass'],'original_dual_guide_accepted':dual['original_upper_inverse_guide_accepted'],
            'full_original_audit':observed,'full_original_dual':dual,'vectors_sha256':sha((folder/'vectors.npz').read_bytes())})
    output={'status':'PASS','optimizer_calls':0,'records':records,'source_data_unchanged':before==frozen_hashes(),
        'scope':'Reassembled original q8 program and independently implemented physical audit; exact frozen references bind old raw digests, structure is exact and coefficient roundoff is explicitly checked. All original physical/dual gates and accept/reject flags are unchanged. No new root or physical target certificate.',
        'program_reference_manifest_sha256':reference_manifest_sha256,
        'comparator_sha256':sha((ROOT/'portable_program_audit.py').read_bytes()),
        'auditor_sha256':sha(Path(__file__).read_bytes()),'producer_summary_sha256':sha((root/'summary.json').read_bytes())}
    write(args.output,output);print(json.dumps({'status':'PASS','records':len(records),'optimizer_calls':0,'output':str(args.output)}))
if __name__=='__main__':main()
