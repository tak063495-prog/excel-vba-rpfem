"""Check the actual CI coefficient artifacts and record the bounded repair."""
from __future__ import annotations
import copy
import json
from pathlib import Path
import sys
sys.dont_write_bytecode=True
from run_fixed_point_experiment import ROOT,core,dual_metrics,frozen_hashes,load_model,program_digest,sha,write
from portable_program_audit import compare_program,compare_metric,unpack_program
from VERIFY_MICRO import load_references,bind_recorded_program_digest

target=ROOT/'reports/portability/CI_PORTABILITY_ANALYSIS.json'
if target.exists():raise FileExistsError(target)
refs,reference_manifest_sha=load_references()
ciroot=ROOT/'reports/ci_assembly';environment=json.loads((ciroot/'micro-programs-environment.json').read_text(encoding='utf-8'))
before=frozen_hashes()
assert environment['source_hashes']==before
observed_programs={};program_comparisons={}
for case in ('P02A','D05A'):
    p=ciroot/f'micro-programs-{case}.npz'
    with core.np.load(p,allow_pickle=False)as data:
        arrays={'E_data':data['0_data'],'E_indices':data['0_indices'],'E_indptr':data['0_indptr'],'E_shape':data['0_shape'],
            'b':data['1'],'G_data':data['2_data'],'G_indices':data['2_indices'],'G_indptr':data['2_indptr'],'G_shape':data['2_shape'],
            'h':data['3'],'dims':data['4'],'cost':data['5']}
        program=unpack_program(arrays)
    assert program_digest(program)==environment['programs'][case]
    observed_programs[case]=program
    comparison=compare_program(program,refs[case][0])
    comparison.update(ci_artifact_path=p.relative_to(ROOT).as_posix(),ci_artifact_sha256=sha(p.read_bytes()),
        actual_CI_digest_matches_environment_exactly=True,
        old_record_binding=bind_recorded_program_digest(refs[case][1]['original_program_sha256'],program,refs[case][0]),
        new_current_host_record_binding=bind_recorded_program_digest(program_digest(program),program,refs[case][0]))
    program_comparisons[case]=comparison
vector_checks=[]
for p in sorted((ROOT/'results/fixed_point_02').glob('*/result.json')):
    meta=json.loads(p.read_text(encoding='utf-8'));mesh,_=load_model(meta['case'])
    pr=core.Prepared(mesh,1.,'upper',8);pr.original=observed_programs[meta['case']]
    with core.np.load(p.parent/'vectors.npz',allow_pickle=False)as data:x,z=data['x'],data['original_dual']
    audit=core.evaluate(pr,x,z);expected=meta['result']['audit'];metrics=[]
    for key in ('physical_pass','witness','numerical_gate'):assert audit[key]is expected[key],(p.parent.name,key)
    for key in ('objective','value','eq_error','cone_error','dual_objective','dual_stationarity','dual_cone_error','dual_gap'):
        metrics.append(compare_metric(audit[key],expected[key],key))
    dual=dual_metrics(pr,z);old=meta['result']['original_dual']
    assert dual['original_upper_inverse_guide_accepted']is old['original_upper_inverse_guide_accepted']
    for key in ('stationarity','cone_error','objective'):metrics.append(compare_metric(dual[key],old[key],'dual.'+key))
    vector_checks.append({'name':p.parent.name,'all_original_gate_decisions_unchanged':True,
        'physical_pass':audit['physical_pass'],'original_dual_guide_accepted':dual['original_upper_inverse_guide_accepted'],
        'metrics':metrics,'scope':'Actual CI coefficient matrices with saved vectors, re-audited locally; full CI host rerun remains a separate workflow record.'})
cpu_replays={}
for name in ('DEFAULT','SANDYBRIDGE','NEHALEM'):
    p=ROOT/f'reports/portability/{name}_REPLAY.json';data=json.loads(p.read_text(encoding='utf-8'))
    assert data['status']=='PASS'and data['optimizer_calls']==0 and len(data['records'])==6
    cpu_replays[name]={'path':p.relative_to(ROOT).as_posix(),'sha256':sha(p.read_bytes()),
        'status':data['status'],'optimizer_calls':0,
        'raw_program_digests':{r['name']:r['original_program_sha256']for r in data['records']},
        'raw_mismatch_count':sum(not r['coefficient_comparison']['raw_sha256_identical']for r in data['records'])}
tests=ROOT/'reports/portability/NEGATIVE_CONTROLS.json';negative=json.loads(tests.read_text(encoding='utf-8'))
assert negative['status']=='PASS'and negative['test_count']==16
assert before==frozen_hashes()
report={'status':'PASS','optimizer_calls':0,'actual_CI_artifact_source':'micro-replay-assembly-diagnostics from workflow artifact39030547569',
    'actual_CI_environment':environment,'actual_CI_source_and_input_hashes_match':True,
    'environment_file_sha256':sha((ciroot/'micro-programs-environment.json').read_bytes()),
    'program_comparisons':program_comparisons,'actual_CI_coefficient_saved_vector_checks':vector_checks,
    'CPU_dispatch_replays':cpu_replays,'negative_tests':{'path':tests.relative_to(ROOT).as_posix(),'sha256':sha(tests.read_bytes()),'test_count':16,'status':'PASS'},
    'coefficient_guard':'Exact CSC structure/shapes/dtypes/integer cones. Finite float64 arrays allow only32*eps*max(1,maxabs(reference)). Reference bytes/SHA/digest bind every old original record.',
    'audit_value_guard':'Recorded/replayed scalar metric differences allow only32*eps*max(1,abs(recorded metric)); every physical/witness/numerical/dual-sign acceptance flag remains exact.',
    'physical_thresholds_unchanged':{'primal_and_physical':1e-7,'original_dual_stationarity':1e-7,'dual_cone':1e-8,'original_upper_dual_objective_min':1.000003},
    'failure_retained':'Original CI job failed because raw assembly bytes differed, although immutable source/inputs matched. Original producer/core/runs/metadata and first strict verifier remain unchanged; first verifier preserved as reports/VERIFY_MICRO_v1_exact_hash.py.',
    'fresh_run_behavior':'A recorded digest must equal either the exact frozen reference or the current reconstructed program. Unknown third-host digests without coefficient evidence fail. Actual CI program tests exercise both binding paths without creating a new-solve claim.',
    'explanation':'Changing only OPENBLAS_CORETYPE to Sandybridge or Nehalem reproduces actual CI raw program digests. np.linalg.norm of two-coordinate edge vectors changes final bits under BLAS dispatch; derived G/cost coefficients change while sparse structure and original acceptance decisions remain identical.',
    'source_data_unchanged':True,'producer_script_unchanged':True,
    'comparator_sha256':sha((ROOT/'portable_program_audit.py').read_bytes()),
    'verifier_sha256':sha((ROOT/'VERIFY_MICRO.py').read_bytes()),
    'program_reference_manifest_sha256':reference_manifest_sha,
    'analysis_script_sha256':sha(Path(__file__).read_bytes())}
write(target,report)
print(json.dumps({'status':'PASS','programs':2,'saved_vectors':len(vector_checks),'CPU_replays':3,'negative_tests':16,'optimizer_calls':0,'report_sha256':sha(target.read_bytes())}))
