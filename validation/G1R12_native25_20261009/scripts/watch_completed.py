"""Postprocess closed native runs while their peers keep solving.

Reads evidence only, writes per-case verification/graded summaries, and never
changes native books, raw statuses, solver profiles, or running Excel objects.
"""
import json,time,csv,traceback
from pathlib import Path
from datetime import datetime,timezone
from prepare_inputs import ROOT
from report_progress import collect,report
from case_locations import case_folder
from check_saved_output import check as check_output
from check_saved_sources import check as check_sources

def tick():
    rows=collect();new=[];errors=[]
    for r in rows:
        if r['status'] in ['RUNNING','PREPARING','NOT_RUN']:continue
        folder=case_folder(r['case'])
        # Allow the worker's immediate JSON/save/close bookkeeping to finish.
        if time.time()-(folder/'status.json').stat().st_mtime<10:continue
        try:
            if r.get('workbook') and not (folder/'saved_sources_check.json').exists():
                check_sources(r['case']);new.append(r['case']+' source')
            if r.get('fs_obtained') and not (folder/'saved_output_check.json').exists():
                check_output(r['case']);new.append(r['case']+' saved output')
        except Exception:
            error={'case':r['case'],'error':traceback.format_exc()};errors.append(error)
            (folder/'postprocess_error.json').write_text(json.dumps(error,ensure_ascii=False,indent=2),encoding='utf-8')
    (ROOT/'results/graded_cases.json').write_text(json.dumps(rows,ensure_ascii=False,indent=2),encoding='utf-8')
    keys=['case','status','native_worker_status','native_run_started','fs_obtained','target_met','lower_fs','upper_fs','relative_gap','root_width','root_search_incomplete','witness_q','solve_seconds','seconds','error','evidence_root']
    temp=ROOT/'results/graded_results.tmp'
    with temp.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.DictWriter(f,fieldnames=keys,extrasaction='ignore');w.writeheader();w.writerows(rows)
    temp.replace(ROOT/'results/graded_results.csv')
    if new:print('VERIFIED',', '.join(new),flush=True)
    if errors:print('POSTPROCESS_ERRORS',json.dumps(errors,ensure_ascii=False),flush=True)
    report()
    terminal=len(rows)==25 and all(r['status'] not in ['RUNNING','PREPARING','NOT_RUN'] for r in rows)
    if terminal:
        checks=all((case_folder(r['case'])/'saved_sources_check.json').exists() for r in rows if r.get('workbook')) and all((case_folder(r['case'])/'saved_output_check.json').exists() for r in rows if r.get('fs_obtained'))
        (ROOT/'results/native25_terminal.json').write_text(json.dumps({'updated_utc':datetime.now(timezone.utc).isoformat(),'all25_native_started':all(r['native_run_started'] for r in rows),'all25_terminal':True,'postprocess_checks_complete':checks,'postprocess_errors':errors},ensure_ascii=False,indent=2),encoding='utf-8')
        if checks:
            print('ACTUAL25_NATIVE_TERMINAL',flush=True);return True
    return False

if __name__=='__main__':
    while not tick():time.sleep(30)
