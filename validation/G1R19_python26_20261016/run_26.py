"""Bounded fresh Python reference solves, with every one of the 26 cases reported.

This tests attainability with the original G1R19 physical equations and an
explicit Python adaptation policy. It does not impersonate Excel's legacy
control flow, import old witnesses, or count historical replays as cold runs.
"""
from pathlib import Path
import argparse
import csv
from datetime import datetime,timezone
import hashlib
import json
import os
import subprocess
import sys
import time
import psutil

ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[1]
TABLE=REPO/'docs/benchmarks/G1R17_native26_20261013/data/RESULTS.csv'

def sha(path):
    h=hashlib.sha256()
    with path.open('rb') as f:
        for data in iter(lambda:f.read(1024*1024),b''):h.update(data)
    return h.hexdigest()

def atomic(path,value):
    temp=path.with_suffix(path.suffix+'.tmp')
    temp.write_text(json.dumps(value,indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8');temp.replace(path)

def source_hashes():
    paths=list((ROOT/'python').rglob('*.py'))+[ROOT/'run_26.py',ROOT/'run_case.py',ROOT/'audit_case.py',ROOT/'INPUTS_26.json']
    paths+=list((REPO/'versions/G1R19/src/vba').glob('*.bas'))
    paths+=[REPO/'versions/G1R19/delivery/rpfem_native_x64.dll',REPO/'versions/G1R19_Issue11/src/vba/RPX_Adapt.bas']
    return {p.relative_to(REPO).as_posix():sha(p) for p in sorted(paths)}

def stop_owned(active):
    process=active['process']
    if process.poll() is not None:return
    candidate=psutil.Process(process.pid)
    if abs(candidate.create_time()-active['create_time'])>.01 or candidate.cmdline()!=active['command']:
        raise RuntimeError('Owned-process identity changed; refusing termination')
    candidate.terminate()
    try:process.wait(timeout=5)
    except subprocess.TimeoutExpired:candidate.kill();process.wait(timeout=5)

def save_progress(output,records,provenance):
    atomic(output/'PROGRESS.json',{'provenance':provenance,'cases':records})
    columns=['case','status','returncode','seconds','lower_fs','upper_fs','gap_percent','physical_gap_pass','root_complete','target_met','lower_rounded_2dp','upper_rounded_2dp','rounded_2dp_agree','driver_budget_expired','audit_error']
    with (output/'RESULTS.csv').open('w',encoding='utf-8-sig',newline='') as f:
        writer=csv.DictWriter(f,fieldnames=columns,extrasaction='ignore');writer.writeheader();writer.writerows(records)

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--output',type=Path,required=True);ap.add_argument('--parallel',type=int,default=2)
    ap.add_argument('--cases',nargs='+');ap.add_argument('--case-seconds',type=float,default=900);ap.add_argument('--solve-seconds',type=float,default=30)
    ap.add_argument('--target-gap',type=float,default=.02);ap.add_argument('--cycles',type=int,default=8);ap.add_argument('--max-elements',type=int,default=32000)
    args=ap.parse_args()
    if not 1<=args.parallel<=2 or args.case_seconds<=0 or args.solve_seconds<=0 or not 0<args.target_gap<=.02:
        raise ValueError('Explicit precision profile and bounded concurrency required')
    all_cases=[r['case'] for r in csv.DictReader(TABLE.open(encoding='utf-8-sig',newline=''))]
    assert len(all_cases)==len(set(all_cases))==26
    cases=args.cases or all_cases
    if len(cases)!=len(set(cases)) or any(c not in all_cases for c in cases):raise ValueError('Unknown or duplicate case')
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    environment=dict(os.environ)
    for key in ('OMP_NUM_THREADS','OPENBLAS_NUM_THREADS','MKL_NUM_THREADS','RAYON_NUM_THREADS','BLIS_NUM_THREADS','VECLIB_MAXIMUM_THREADS'):
        environment[key]='1'
    sources=source_hashes()
    provenance={'scope':'Fresh Python reference solve; no seed/witness import; q8 original physical gates; not an Excel timing test',
        'started_utc':datetime.now(timezone.utc).isoformat(),'source_sha256':sources,'expected_cases':cases,'total_catalog_cases':26,
        'target_gap':args.target_gap,'per_case_seconds':args.case_seconds,'per_solve_seconds':args.solve_seconds,'hard_watchdog_grace_seconds':60,
        'cycles':args.cycles,'max_elements':args.max_elements,'parallel':args.parallel,'solver_threads':1,'fan_lower':True,
        'g1r19_tag':subprocess.check_output(['git','rev-parse','G1R19'],cwd=REPO,text=True).strip(),
        'implementation_commit':subprocess.check_output(['git','rev-parse','HEAD'],cwd=REPO,text=True).strip(),
        'catalog_sha256':sha(TABLE),'no_historical_pairs_counted_as_cold':True}
    records=[{'case':c,'status':'PENDING','target_met':False} for c in cases]
    pending=list(records);active={};save_progress(output,records,provenance)
    flags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0
    try:
        while pending or active:
            if pending and len(active)<args.parallel and source_hashes()!=sources:
                raise RuntimeError('Frozen validation sources changed during the batch')
            while pending and len(active)<args.parallel:
                record=pending.pop(0);case=record['case']
                command=[sys.executable,'-X','utf8',str(ROOT/'run_case.py'),'--case',case,'--output',str(output/'cases'),
                    '--target-gap',str(args.target_gap),'--case-seconds',str(args.case_seconds),'--solve-seconds',str(args.solve_seconds),
                    '--cycles',str(args.cycles),'--max-elements',str(args.max_elements),'--fan-lower']
                log=(output/f'{case}.log').open('x',encoding='utf8')
                process=subprocess.Popen(command,cwd=REPO,env=environment,stdout=log,stderr=subprocess.STDOUT,stdin=subprocess.DEVNULL,creationflags=flags)
                created=psutil.Process(process.pid).create_time()
                record.update(status='RUNNING',pid=process.pid,process_create_time=created,command=command,started_utc=datetime.now(timezone.utc).isoformat())
                active[case]={'process':process,'create_time':created,'command':command,'record':record,'log':log,'start':time.monotonic(),'forced_timeout':False}
                print('START',case,process.pid,flush=True);save_progress(output,records,provenance)
            finished=[]
            for case,current in active.items():
                process=current['process'];elapsed=time.monotonic()-current['start']
                if process.poll() is None and elapsed>args.case_seconds+60:
                    stop_owned(current);current['forced_timeout']=True
                if process.poll() is None:continue
                current['log'].close();record=current['record'];record.update(returncode=process.returncode,seconds=elapsed,
                    finished_utc=datetime.now(timezone.utc).isoformat(),forced_timeout=current['forced_timeout'])
                summary_path=output/'cases'/case/'summary.json'
                try:summary=json.loads(summary_path.read_text(encoding='utf8')) if summary_path.exists() else {}
                except (OSError,ValueError):summary={}
                audit_path=output/f'{case}_independent_audit.json'
                audit_command=[sys.executable,'-X','utf8',str(ROOT/'audit_case.py'),'--case',case,'--folder',str(output/'cases'/case),
                    '--target-gap',str(args.target_gap),'--output',str(audit_path)]
                try:
                    audit_run=subprocess.run(audit_command,cwd=REPO,env=environment,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,
                        encoding='utf8',timeout=120,creationflags=flags)
                    (output/f'{case}_independent_audit.log').write_text(audit_run.stdout,encoding='utf8')
                    evidence=json.loads(audit_path.read_text(encoding='utf8')) if audit_run.returncode==0 else {}
                    if audit_run.returncode!=0:record['audit_error']=audit_run.stdout[-2000:]
                except subprocess.TimeoutExpired:record['audit_error']='INDEPENDENT_AUDIT_TIMEOUT';evidence={}
                record.update(lower_fs=evidence.get('lower_fs'),upper_fs=evidence.get('upper_fs'),
                    gap_percent=None if evidence.get('relative_gap') is None else 100*evidence['relative_gap'],
                    lower_rounded_2dp=evidence.get('lower_rounded_2dp'),upper_rounded_2dp=evidence.get('upper_rounded_2dp'),rounded_2dp_agree=evidence.get('rounded_2dp_agree'),
                    physical_gap_pass=evidence.get('physical_gap_pass',False),root_complete=evidence.get('root_complete',False),
                    target_met=evidence.get('target_met',False),driver_budget_expired=bool(summary.get('budget_exhausted') or summary.get('case_budget',{}).get('exhausted')))
                record['status']=('PASS' if evidence.get('target_met') and process.returncode==0 and summary.get('target_met') else
                    'VERIFIED_PAIR_AFTER_INTERRUPTION' if evidence.get('target_met') else
                    'BUDGET_EXPIRED' if current['forced_timeout'] or record['driver_budget_expired'] else
                    'ERROR' if process.returncode!=0 or record.get('audit_error') else evidence.get('status','NO_CERTIFIED_PAIR'))
                case_folder=output/'cases'/case
                record['retained_files_sha256']={p.relative_to(case_folder).as_posix():sha(p) for p in sorted(case_folder.rglob('*')) if p.is_file()}
                print('FINISH',case,record['status'],record.get('gap_percent'),round(elapsed,2),flush=True)
                finished.append(case);save_progress(output,records,provenance)
            for case in finished:del active[case]
            if active:time.sleep(1)
    finally:
        for current in active.values():stop_owned(current);current['log'].close()
        provenance['finished_utc']=datetime.now(timezone.utc).isoformat()
        provenance['sources_unchanged']=source_hashes()==sources
        save_progress(output,records,provenance)
    assert provenance['sources_unchanged']
    result={'scope':provenance['scope'],'cases':len(records),'passed':sum(r['status']=='PASS' for r in records),
        'verified_pair_2pct':sum(r.get('target_met',False) for r in records),'failed_or_incomplete':[r['case'] for r in records if r['status']!='PASS']}
    atomic(output/'SUMMARY.json',result);print(json.dumps(result),flush=True)

if __name__=='__main__':main()
