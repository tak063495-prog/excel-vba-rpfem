"""Exact-owned Windows numeric-tree watchdog, including redirector children."""
from __future__ import annotations
import argparse
import hashlib
import json
import os
from pathlib import Path
import subprocess
import sys
import time
import psutil

ROOT=Path(__file__).resolve().parent
def write(path,obj):
    temp=path.with_name(path.name+'.partial');temp.write_text(json.dumps(obj,ensure_ascii=False,indent=2)+'\n',encoding='utf-8');temp.replace(path)
def sha(path):
    h=hashlib.sha256()
    with Path(path).open('rb') as f:
        for b in iter(lambda:f.read(1048576),b''):h.update(b)
    return h.hexdigest()
def identity(pid):
    p=psutil.Process(pid);return dict(pid=pid,create_time=p.create_time(),cmdline=p.cmdline(),exe=p.exe())
def same(r):
    try:return identity(r['pid'])==r
    except (psutil.NoSuchProcess,psutil.AccessDenied):return False
def discover(job):
    if not same(job['parent']):return
    for p in psutil.Process(job['parent']['pid']).children(recursive=True):
        try:r=identity(p.pid)
        except (psutil.NoSuchProcess,psutil.AccessDenied):continue
        if any(Path(x).name==job['script_name'] for x in r['cmdline']):job['descendants'][str(r['pid'])]=r
def stop(job,reason):
    discover(job);folder=Path(job['folder'])
    before={str(p.relative_to(folder)):sha(p) for p in folder.rglob('*') if p.is_file() and p.suffix in ('.npz','.json')}
    stopped=[]
    for r in list(reversed(list(job['descendants'].values())))+[job['parent']]:
        if same(r):psutil.Process(r['pid']).kill();stopped.append(r)
    after={str(p.relative_to(folder)):sha(p) for p in folder.rglob('*') if p.is_file() and p.suffix in ('.npz','.json')}
    job['interruption']=dict(reason=reason,stopped=stopped,checkpoint_sha256_before=before,checkpoint_sha256_after=after,new_bound_adopted=False)

def run(args):
    specs=json.loads(args.jobs.read_text(encoding='utf-8'));args.output.mkdir(parents=True,exist_ok=False)
    files=list((ROOT/'python').rglob('*.py'))+list(ROOT.glob('*.py'))+[ROOT/'INPUTS_26.json',args.jobs]
    frozen={str(p.resolve()):sha(p) for p in files if '__pycache__' not in p.parts}
    state=dict(status='RUNNING',started_epoch=time.time(),supervisor=identity(os.getpid()),source_sha256=frozen,jobs={})
    env=dict(os.environ,OMP_NUM_THREADS='1',OPENBLAS_NUM_THREADS='1',MKL_NUM_THREADS='1',NUMEXPR_NUM_THREADS='1')
    active={}
    try:
        for spec in specs:
            label=spec['label'];folder=(args.output/label).resolve()
            argv=[x.replace('{OUTPUT}',str(folder)) for x in spec['argv']]
            command=[str(args.python),'-X','utf8',str((ROOT/spec['script']).resolve())]+argv
            log=(args.output/(label+'.log')).open('wb',buffering=0);launched=time.time()
            p=subprocess.Popen(command,stdout=log,stderr=subprocess.STDOUT,cwd=ROOT.parents[1],env=env,creationflags=subprocess.CREATE_NO_WINDOW)
            job=dict(label=label,case=spec['case'],mode=spec['mode'],phase='SOLVING',folder=str(folder),
                script_name=spec['script'],command=command,parent=identity(p.pid),descendants={},started_epoch=launched,
                deadline_epoch=launched+spec['case_seconds']+60,case_seconds=spec['case_seconds'],threads=1,returncode=None)
            state['jobs'][label]=job;active[label]=(p,log)
        while active:
            for label,(p,log) in list(active.items()):
                job=state['jobs'][label];discover(job)
                if p.poll() is None and time.time()>job['deadline_epoch']:stop(job,'candidate whole wall deadline');p.wait(timeout=15)
                if p.poll() is not None:
                    job.update(phase='SOLVER_EXITED',returncode=p.returncode,finished_epoch=time.time())
                    if any(same(r) for r in job['descendants'].values()):stop(job,'surviving numeric child')
                    log.close();del active[label]
            write(args.output/'OWNERSHIP.json',state)
            if active:time.sleep(.5)
        for label,job in state['jobs'].items():
            command=[str(args.python),'-X','utf8',str(ROOT/'audit_side.py'),'--case',job['case'],'--mode',job['mode'],
                '--folder',job['folder'],'--output',str(args.output/(label+'_INDEPENDENT_AUDIT.json'))]
            with (args.output/(label+'_AUDIT.log')).open('wb') as f:
                p=subprocess.Popen(command,stdout=f,stderr=subprocess.STDOUT,env=env,creationflags=subprocess.CREATE_NO_WINDOW)
                ajob=dict(parent=identity(p.pid),descendants={},script_name='audit_side.py',folder=job['folder'])
                job.update(phase='AUDITING',audit_owner=ajob['parent']);write(args.output/'OWNERSHIP.json',state)
                until=time.time()+90
                while p.poll() is None and time.time()<until:discover(ajob);time.sleep(.2)
                if p.poll() is None:stop(ajob,'independent no-solver audit deadline')
                job.update(audit_returncode=p.wait(timeout=15),phase='FINISHED')
                if any(same(r) for r in ajob['descendants'].values()):stop(ajob,'surviving audit child')
        state.update(status='FINISHED',finished_epoch=time.time(),source_sha256_after={p:sha(p) for p in frozen})
        state['source_unchanged']=state['source_sha256_after']==frozen
        if not state['source_unchanged']:state['status']='SOURCE_CHANGED'
        write(args.output/'OWNERSHIP.json',state)
        print('SIDES_FINISHED',state['status'],state['source_unchanged'],flush=True)
    finally:
        for label,(p,log) in active.items():
            if p.poll() is None:stop(state['jobs'][label],'supervisor cleanup')
            log.close()

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--jobs',type=Path,required=True);p.add_argument('--output',type=Path,required=True)
    p.add_argument('--python',type=Path,default=Path(sys.executable));run(p.parse_args())
