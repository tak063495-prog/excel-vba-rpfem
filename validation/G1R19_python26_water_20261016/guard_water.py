"""One owned, fresh D08A solve with a600-second process-tree watchdog."""
from pathlib import Path
import argparse,json,os,subprocess,sys,time
from datetime import datetime,timezone
import psutil

ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[1]


def write(path,value):
    temporary=path.with_suffix('.json.tmp')
    temporary.write_text(json.dumps(value,indent=2,ensure_ascii=False)+'\n',encoding='utf8');temporary.replace(path)


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--seconds',type=float,default=600.)
    args=parser.parse_args()
    if not 0<args.seconds<=600:raise ValueError('A finite wall limit at most600seconds is required')
    output=args.output.resolve();output.mkdir(parents=True,exist_ok=False)
    env=dict(os.environ)
    for key in ('OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS','BLIS_NUM_THREADS','RAYON_NUM_THREADS'):
        env[key]='1'
    command=[sys.executable,'-X','utf8',str(ROOT/'run_case.py'),'--case','D08A','--output',str(output/'cases'),
             '--target-gap','.02','--case-seconds',str(args.seconds),'--solve-seconds','30',
             '--cycles','8','--max-elements','32000','--fan-lower']
    flags=subprocess.CREATE_NO_WINDOW if os.name=='nt' else 0
    log=(output/'D08A.log').open('x',encoding='utf8')
    started=time.monotonic();process=subprocess.Popen(command,cwd=REPO,env=env,stdout=log,stderr=subprocess.STDOUT,
                                                     stdin=subprocess.DEVNULL,creationflags=flags)
    parent=psutil.Process(process.pid)
    owner={'pid':process.pid,'create_time':parent.create_time(),'command':parent.cmdline()}
    known={};record={'started_utc':datetime.now(timezone.utc).isoformat(),'requested_command':command,
                    'owner':owner,'hard_limit_seconds':args.seconds,'owned_tree':[],'timeout':False}

    def observe():
        candidate=psutil.Process(owner['pid'])
        if candidate.create_time()!=owner['create_time'] or candidate.cmdline()!=owner['command']:
            raise RuntimeError('Owned worker identity changed')
        current=[candidate]+candidate.children(recursive=True)
        for member in current:
            identity={'pid':member.pid,'create_time':member.create_time(),'command':member.cmdline(),
                      'parent_pid':member.ppid()}
            if member.pid in known and known[member.pid]!=identity:
                raise RuntimeError('A retained child process changed identity')
            known[member.pid]=identity
        record['owned_tree']=list(known.values());record['elapsed_seconds']=time.monotonic()-started
        write(output/'OWNERSHIP.json',record)
        return current

    print('D08A_START',json.dumps(owner),flush=True)
    try:
        while process.poll() is None:
            current=observe()
            if time.monotonic()-started>=args.seconds:
                record['timeout']=True
                # Terminate the current verified descendants before the Windows
                # venv redirector. Killing only the redirector leaves its worker.
                for member in reversed(current):
                    identity=known[member.pid]
                    try:
                        if member.create_time()!=identity['create_time'] or member.cmdline()!=identity['command']:
                            raise RuntimeError('Refusing to stop a process with changed identity')
                        member.kill()
                    except psutil.NoSuchProcess:pass
                process.wait(timeout=5)
                break
            time.sleep(.5)
    finally:
        log.close()
        record.update(returncode=process.poll(),elapsed_seconds=time.monotonic()-started,
                      finished_utc=datetime.now(timezone.utc).isoformat())
        write(output/'OWNERSHIP.json',record)
    audit_command=[sys.executable,'-X','utf8',str(ROOT/'audit_case.py'),'--case','D08A',
                   '--folder',str(output/'cases/D08A'),'--output',str(output/'INDEPENDENT_AUDIT.json')]
    completed=subprocess.run(audit_command,cwd=REPO,env=env,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,
                             text=True,encoding='utf8',timeout=120,creationflags=flags)
    (output/'INDEPENDENT_AUDIT.log').write_text(completed.stdout,encoding='utf8')
    print('D08A_DONE',json.dumps(record),flush=True)
    print(completed.stdout,flush=True)


if __name__=='__main__':main()
