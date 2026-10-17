"""Produce the short, explicit-scope Issue28 diagnostic report."""
import json
from pathlib import Path
import sys
sys.dont_write_bytecode=True
from run_fixed_point_experiment import ROOT,write,sha

run=ROOT/'results/fixed_point_02'
summary=json.loads((run/'summary.json').read_text(encoding='utf-8'))
rows=[]
for entry in summary['runs']:
    meta=json.loads((run/entry['name']/'result.json').read_text(encoding='utf-8'))
    r=meta['result'];a=r['audit'];dual=r['original_dual'];errors=a['audit']['errors']
    rows.append({'name':entry['name'],'solver_status':r['status'],'iterations':r['iterations'],
        'solver_seconds':r['solver_seconds'],'original_objective':a['objective'],
        'physical_pass':a['physical_pass'],'original_work_error':errors['work'],
        'original_eq_error':a['eq_error'],'original_cone_error':a['cone_error'],
        'original_dual_stationarity':dual['stationarity'],'original_dual_cone_error':dual['cone_error'],
        'original_dual_objective':dual['objective'],'original_dual_guide_accepted':dual['original_upper_inverse_guide_accepted'],
        'primal_witness':a['witness'],'root_complete':False,'process_exit_code':entry['exit_code'],
        'result_path':str((run/entry['name']/'result.json').relative_to(ROOT)),
        'result_sha256':sha((run/entry['name']/'result.json').read_bytes())})
first=json.loads((ROOT/'results/fixed_point_01/summary.json').read_text(encoding='utf-8'))
scaled=json.loads((run/'D05A_none/result.json').read_text(encoding='utf-8'))
zero=scaled['original_program']['zero_original_dual']
report={'status':'DIAGNOSTIC_EXPERIMENT_COMPLETE','scope':summary['scope'],'new_optimizer_calls_final_run':6,
    'initial_run_optimizer_calls':first['optimizer_calls'],'new_optimizer_calls_total_including_harness_failures':first['optimizer_calls']+6,
    'final_run_elapsed_seconds':summary['process_elapsed_seconds'],'per_solve_time_limit_seconds':20.,
    'geometry_elements':214,'fs':1.,'q':8,'rows':rows,'similarity_law':summary['similarity_law'],
    'zero_dual_analysis':{'D05A_raw_stationarity':zero['stationarity'],
        'normalized_cost_stationarity':scaled['original_program']['zero_candidate_normalized_objective_stationarity'],
        'relative_1e_minus_7_times_maxcost_limit':1e-7*scaled['original_program']['VBA_style_QScale'],
        'relative_gate_would_accept_zero':False,'original_dual_objective':zero['objective'],
        'required_original_dual_objective_min':1.000003,
        'conclusion':'A near-zero dual cannot be made useful just by replacing 1e-7 with 1e-7*max(1,max(abs(cost))). Zero gives normalized stationarity=1 and objective=0, both reject.'},
    'inference':'Coordinate/cone conditioning is a supported experimental suspect; these Clarabel LPs do not identify a bug in the unavailable current 400-element HSD/helper vectors.',
    'negative_control':'Normalized-objective route obtains a valid original dual sign guide but its primal physical work error is above 1e-7; do not publish that primal as a bound.',
    'initial_harness_failure':{'run':'results/fixed_point_01','status':first['status'],
        'failed_children':[r['name']for r in first['runs']if r['exit_code']!=0],
        'cause':'New wrapper interpreted unchanged experiment_core.solve third return as restored original dual. It is canonical reduced dual. Explicit pr.restore_dual fixes wrapper; core/source/gates unchanged.',
        'preserved_original_script':'reports/harness_initial_v1.py','preserved_script_sha256':sha((ROOT/'reports/harness_initial_v1.py').read_bytes()),
        'final_script_sha256':sha((ROOT/'run_fixed_point_experiment.py').read_bytes())},
    'source_data_unchanged':summary['source_data_unchanged'],'native_or_VBA_executed':False,'new_root_or_gap_target_certified':False,
    'new_vector_replay_report':'reports/MICRO_REPLAY_RESULT.json','new_vector_replay_status':'PASS','new_vector_replay_optimizer_calls':0,
    'published_report_vs_actual_reproduction':'Issue28 reports 400-element Fs1 objective2.40457/raw_stat1429.67127988363. This new 214-element diagnostic has original objective near2.583159/maxcost1906.2283737862776. It verifies the scale mechanism, not those exact runtime numbers.'}
write(ROOT/'reports/MICRO_EXPERIMENT_SUMMARY.json',report)
table='| Case / route | status | original objective | work error | physical | original dual stat | original dual guide |\n|---|---|---:|---:|---|---:|---|\n'
for row in rows:
    table+=f"| {row['name']} | {row['solver_status']} | {row['original_objective']:.12g} | {row['original_work_error']:.5g} | {row['physical_pass']} | {row['original_dual_stationarity']:.5g} | {row['original_dual_guide_accepted']} |\n"
text=f"""# Issue28: new bounded Python conditioning experiment\n\nSix new upper-LP solves, q=8 and Fs=1, on the recorded same-topology 214-element P02A / D05A geometry. Each solver was limited to20 seconds and1 thread; final six child processes exited normally in {summary['process_elapsed_seconds']:.3f} seconds total. This duration is for a small diagnostic only.\n\n{table}\nAll results, raw vectors, source/input hashes and process exit records are under `results/fixed_point_02`. `VERIFY_MICRO.py` reassembles the original programs and replays all six vectors with zero optimization calls; it matches both passes and intentional rejections.\n\n## What is verified\n\nP02A is the physical source of D05A: coordinates x1000, gamma /1000, cohesion/phi/psi unchanged. The recorded coarse mesh connectivity is identical. Cost grows1000x, strain rows shrink1000x, reference-work RHS shrinks1000x, and jump rows remain the same, up to recorded floating-point errors. Exact positive primal and dual maps preserve the original objective. No physical material/load parameter or gate was changed.\n\nThe ordinary original and affine D05A attempts fail the original work/physical gates. A pure objective normalization gives a valid original dual sign guide after multiplying its dual by QScale, but the primal still has work error1.6672e-6 and is rejected. The exact positive velocity-coordinate substitution plus a common positive multiplier per cone yields physical and raw original-dual passes with work error6.66e-16 and stationarity6.1404e-8. Solver status is AlmostSolved; the pass comes from original audits, not status.\n\nAt Fs=1 the physical objective is about2.583159, above the upper witness threshold. Therefore every `witness` flag is false. This is a fixed-point inverse-sign diagnostic; it establishes neither an upper Fs bound nor roots nor GAP2% completion.\n\nZero dual has raw stationarity1906.2283737862776, normalized stationarity1 and original objective0. A relative1e-7 tolerance still rejects it by seven orders of magnitude and its objective is below1.000003. Thus relaxing the raw stationarity threshold alone cannot rescue the issue's reported near-zero candidate. A dual solved with normalized cost must be converted to the original objective's units before original stationarity/objective gates. The intentionally unconverted diagnostic is stored as rejected.\n\n## Limits and retained failure\n\nThe current Issue28 raw400-element HSD/native/helper vectors were unavailable. This214-element Clarabel experiment does not reproduce the reported exact1429.671 residual, identify a DLL defect, or verify a VBA fix. It supports conditioning/unit restoration as a direction to investigate. Product source and prior results are untouched.\n\n`results/fixed_point_01` is retained with two wrapper errors: the new harness initially treated the unchanged core's third return as original dual, although that API returns reduced dual. The first harness is preserved as `reports/harness_initial_v1.py`. Only the new wrapper was corrected to explicitly restore original duals, followed by the complete second run. All12 attempted optimizer calls across both runs are disclosed; the numerical core and gates were unchanged.\n"""
(ROOT/'reports/MICRO_EXPERIMENT_SUMMARY.md').write_text(text,encoding='utf-8')
print(json.dumps({'status':report['status'],'rows':len(rows),'total_optimizer_calls':report['new_optimizer_calls_total_including_harness_failures'],'report':'reports/MICRO_EXPERIMENT_SUMMARY.json'}))
