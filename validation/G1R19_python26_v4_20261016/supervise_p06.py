"""Owned Windows parent/solver-child supervision and bounded independent audit."""
from __future__ import annotations
import argparse
from datetime import datetime, timezone
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import traceback
import psutil

ROOT=Path(__file__).resolve().parent


def write(path,value):
    path=Path(path);tmp=path.with_name(path.name+'.partial')
    tmp.write_text(json.dumps(value,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');tmp.replace(path)


def sha(path):
    h=hashlib.sha256()
    with Path(path).open('rb') as stream:
        for chunk in iter(lambda:stream.read(1024*1024),b''):h.update(chunk)
    return h.hexdigest()


def identity(pid):
    p=psutil.Process(pid)
    return dict(pid=pid,create_time=p.create_time(),cmdline=p.cmdline(),exe=p.exe())


def same_process(record):
    try:
        p=psutil.Process(record['pid'])
        return bool(abs(p.create_time()-record['create_time'])<1e-6 and p.cmdline()==record['cmdline'])
    except (psutil.NoSuchProcess,psutil.AccessDenied):return False


def discover(job):
    if not same_process(job['parent']):return
    try:children=psutil.Process(job['parent']['pid']).children(recursive=True)
    except (psutil.NoSuchProcess,psutil.AccessDenied):return
    for p in children:
        try:r=identity(p.pid)
        except (psutil.NoSuchProcess,psutil.AccessDenied):continue
        cmd=r['cmdline']
        if ('--case' in cmd and cmd[cmd.index('--case')+1]==job['case']
                and any(Path(x).name=='run_case.py' for x in cmd)):
            job['descendants'][str(p.pid)]=r


def checkpoint_hashes(folder):
    return {str(p.relative_to(folder)):sha(p) for p in folder.rglob('*')
            if p.is_file() and p.suffix in ('.npz','.json') and not p.name.endswith('.partial')}


def terminate_owned(job,reason):
    discover(job);folder=Path(job['folder'])
    before=checkpoint_hashes(folder) if folder.exists() else {}
    stopped=[]
    # The venv redirector is a real parent. Kill only exact recorded numeric
    # descendants, then that exact parent; never a broad Python process list.
    for r in reversed(list(job['descendants'].values()))+[job['parent']]:
        if same_process(r):
            psutil.Process(r['pid']).kill();stopped.append(r)
    after=checkpoint_hashes(folder) if folder.exists() else {}
    evidence=dict(reason=reason,stopped=stopped,checkpoint_sha256_before=before,
                  checkpoint_sha256_after=after,new_bound_adopted=False,
                  wall_epoch=time.time(),scope='owned PID/create_time/full-command tree only')
    write(Path(job['ownership_path']).with_name(job['case']+'_INTERRUPTION_'+str(time.time_ns())+'.json'),evidence)
    job['interruption']=evidence


def run(args):
    args.output.mkdir(parents=True,exist_ok=False)
    files=sorted((ROOT/'python').rglob('*.py'))+[ROOT/'run_case.py',Path(__file__),ROOT/'audit_case.py',ROOT/'INPUTS_26.json']
    frozen={str(p.resolve()):sha(p) for p in files if '__pycache__' not in p.parts}
    write(args.output/'SOURCE_FREEZE.json',frozen)
    state=dict(status='RUNNING',started_utc=datetime.now(timezone.utc).isoformat(),
               supervisor=identity(os.getpid()),source_sha256=frozen,cases={})
    write(args.output/'OWNERSHIP.json',state)
    active={};streams=[];audit_live=None
    env=dict(os.environ)
    for key in ('OMP_NUM_THREADS','OPENBLAS_NUM_THREADS','MKL_NUM_THREADS','NUMEXPR_NUM_THREADS'):env[key]='1'
    try:
        for case in ('P06A','P06N'):
            native=args.native_mesh/case/'target_1000_mesh.npz'
            cmd=[str(args.python),'-X','utf8',str(ROOT/'run_case.py'),'--case',case,
                '--initial-mesh',str(native.resolve()),'--output',str((args.output/'cases').resolve()),
                '--cycles','8','--max-elements','32000','--theta','.8','--lower-theta','.8',
                '--fan-lower','--upper-indicator','mixed','--kinematic-fraction','.2',
                '--root-method','bisection','--initial-span','.005','--target-gap','.02',
                '--case-seconds',str(args.case_seconds),'--solve-seconds',str(args.solve_seconds)]
            log=(args.output/(case+'.log')).open('wb',buffering=0);streams.append(log)
            launched=time.time()
            process=subprocess.Popen(cmd,stdout=log,stderr=subprocess.STDOUT,cwd=ROOT.parents[1],env=env,
                                     creationflags=subprocess.CREATE_NO_WINDOW)
            job=dict(case=case,phase='SOLVING',folder=str((args.output/'cases'/case).resolve()),
                command=cmd,parent=identity(process.pid),descendants={},started_epoch=launched,
                deadline_epoch=launched+args.case_seconds+60,case_seconds=args.case_seconds,
                solve_seconds=args.solve_seconds,threads=1,returncode=None,
                ownership_path=str((args.output/(case+'_OWNERSHIP.json')).resolve()))
            state['cases'][case]=job;active[case]=(process,log)
            write(job['ownership_path'],job);write(args.output/'OWNERSHIP.json',state)
        heartbeat=0.
        while active:
            for case,(process,log) in list(active.items()):
                job=state['cases'][case];discover(job)
                if process.poll() is None and time.time()>job['deadline_epoch']:
                    terminate_owned(job,'case wall deadline including redirector grace');process.wait(timeout=15)
                code=process.poll()
                if code is None:continue
                job['returncode']=code
                remaining=[r for r in job['descendants'].values() if same_process(r)]
                if remaining:
                    terminate_owned(job,'unexpected surviving numeric child after parent exit')
                job['finished_epoch']=time.time();job['phase']='SOLVER_EXITED'
                log.close();del active[case]
                write(job['ownership_path'],job);write(args.output/'OWNERSHIP.json',state)
            if time.time()-heartbeat>=2:
                for job in state['cases'].values():write(job['ownership_path'],job)
                write(args.output/'OWNERSHIP.json',state);heartbeat=time.time()
            if active:time.sleep(.5)
        # A fresh process independently reassembles the unchanged original q8
        # program. This audit starts no optimization and is bounded separately.
        for case,job in state['cases'].items():
            out=args.output/(case+'_INDEPENDENT_AUDIT.json')
            command=[str(args.python),'-X','utf8',str(ROOT/'audit_case.py'),'--case',case,
                '--folder',job['folder'],'--target-gap','.02','--output',str(out.resolve())]
            with (args.output/(case+'_AUDIT.log')).open('wb') as stream:
                proc=subprocess.Popen(command,stdout=stream,stderr=subprocess.STDOUT,cwd=ROOT.parents[1],env=env,
                                      creationflags=subprocess.CREATE_NO_WINDOW)
                ajob=dict(case=case,parent=identity(proc.pid),descendants={},folder=job['folder'],
                          ownership_path=job['ownership_path'])
                audit_live=(proc,ajob)
                job['audit_owner']=ajob['parent'];job['audit_command']=command;job['phase']='AUDITING'
                write(job['ownership_path'],job)
                until=time.monotonic()+120
                while proc.poll() is None and time.monotonic()<until:
                    # Audit uses audit_case.py, not run_case.py; record its exact
                    # children with the same case and audit script separately.
                    try:
                        for child in psutil.Process(proc.pid).children(recursive=True):
                            r=identity(child.pid)
                            if any(Path(x).name=='audit_case.py' for x in r['cmdline']):ajob['descendants'][str(child.pid)]=r
                    except (psutil.NoSuchProcess,psutil.AccessDenied):pass
                    time.sleep(.25)
                if proc.poll() is None:terminate_owned(ajob,'independent no-solver audit 120s deadline')
                job['audit_returncode']=proc.wait(timeout=15)
                if any(same_process(r) for r in ajob['descendants'].values()):
                    terminate_owned(ajob,'unexpected surviving independent audit child')
                audit_live=None
            job['phase']='FINISHED';job['independent_audit']=str(out.resolve()) if out.exists() else None
            if out.exists():job['independent_audit_status']=json.loads(out.read_text(encoding='utf-8')).get('status')
            write(job['ownership_path'],job);write(args.output/'OWNERSHIP.json',state)
        after={p:sha(p) for p in frozen};state['source_unchanged']=after==frozen
        state['source_sha256_after']=after;state['status']='FINISHED' if after==frozen else 'SOURCE_CHANGED'
        state['finished_utc']=datetime.now(timezone.utc).isoformat()
        write(args.output/'OWNERSHIP.json',state)
        print(json.dumps({case:{k:job.get(k) for k in ('returncode','audit_returncode','independent_audit_status')}
                          for case,job in state['cases'].items()}),flush=True)
    except BaseException:
        for case,(process,log) in active.items():
            if process.poll() is None:terminate_owned(state['cases'][case],'supervisor exception cleanup')
        if audit_live is not None and audit_live[0].poll() is None:
            terminate_owned(audit_live[1],'supervisor audit exception cleanup')
        state['status']='ERROR';state['error']=traceback.format_exc();write(args.output/'OWNERSHIP.json',state)
        raise
    finally:
        for stream in streams:
            if not stream.closed:stream.close()


if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--output',type=Path,required=True)
    p.add_argument('--native-mesh',type=Path,required=True)
    p.add_argument('--python',type=Path,default=Path(sys.executable))
    p.add_argument('--case-seconds',type=float,default=900.);p.add_argument('--solve-seconds',type=float,default=30.)
    args=p.parse_args()
    if args.case_seconds<=0 or args.solve_seconds<=0:raise ValueError('Positive bounded times required')
    run(args)
