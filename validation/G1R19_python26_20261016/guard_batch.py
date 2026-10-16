"""External wall watchdog, including Windows venv redirector descendants.

This runs independently of the numerical/audit scheduler. Only processes whose
birth time and complete command have been captured from a registered owned
parent are terminated. Numerical children stop before the redirector parent.
"""
from pathlib import Path
import argparse
import hashlib
import json
import time
import psutil
import win32file
import win32con
import pywintypes

def atomic(path,value):
    temp=path.with_suffix('.tmp');temp.write_text(json.dumps(value,indent=2,ensure_ascii=False)+'\n',encoding='utf8')
    for attempt in range(30):
        try:temp.replace(path);return True
        except PermissionError:
            if attempt<29:time.sleep(.02)
    # A transient external reader must not disable the actual deadline guard.
    print('LEDGER_REPLACE_SHARING_CONFLICT; watchdog remains active',flush=True)
    return False

def read(path,default):
    handle=None
    try:
        # Ordinary Windows open() denies deletion while reading and can make
        # the producer's atomic checkpoint replacement fail. Share deletion.
        handle=win32file.CreateFile(str(path),win32con.GENERIC_READ,
            win32con.FILE_SHARE_READ|win32con.FILE_SHARE_WRITE|win32con.FILE_SHARE_DELETE,
            None,win32con.OPEN_EXISTING,win32con.FILE_ATTRIBUTE_NORMAL,None)
        size=win32file.GetFileSize(handle)
        _,data=win32file.ReadFile(handle,size)
        # Release the filesystem handle before decoding/parsing a large log.
        # FILE_SHARE_DELETE alone is not a guarantee on every workspace drive.
        handle.Close();handle=None
        return json.loads(data.decode('utf8'))
    except (OSError,ValueError,pywintypes.error):return default
    finally:
        if handle is not None:handle.Close()

def compatible(actual,expected):
    if not actual or len(actual)!=len(expected) or actual[1:]!=expected[1:]:return False
    return Path(actual[0]).name.lower() in ('python.exe','pythonw.exe','python','python3','python3.13') and Path(expected[0]).name.lower() in ('python.exe','pythonw.exe','python','python3','python3.13')

def capture(process):
    return {'pid':process.pid,'create_time':process.create_time(),'command':process.cmdline()}

def exact(record):
    try:
        process=psutil.Process(record['pid'])
        return process if abs(process.create_time()-record['create_time'])<.01 and process.cmdline()==record['command'] else None
    except (psutil.NoSuchProcess,psutil.AccessDenied):return None

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--output',type=Path,required=True);args=ap.parse_args();output=args.output.resolve()
    previous=read(output/'GUARD_LEDGER.json',{})
    entries=previous.get('entries',{});ledger={'scope':'External exact-owner child-before-parent guard; independent of the auditor scheduler','entries':entries}
    while not (output/'GUARD_STOP').exists():
        progress=read(output/'PROGRESS.json',{});profile=progress.get('provenance',{})
        registrations=[]
        for r in progress.get('cases',[]):
            if r.get('status')=='RUNNING':
                registrations.append({'key':r['case'],'pid':r['pid'],'create_time':r['process_create_time'],
                    'command':r['command'],'deadline':r['process_create_time']+profile.get('per_case_seconds',900)+50,
                    'case_folder':str(output/'cases'/r['case'])})
        registrations+=read(output/'EXTRA_GUARDS.json',[])
        cancellations=read(output/'CANCELLATIONS.json',{})
        for registration in registrations:
            key=registration['key'];entry=entries.get(key)
            if entry is None:
                try:
                    process=psutil.Process(registration['pid'])
                    if abs(process.create_time()-registration['create_time'])>=.01 or not compatible(process.cmdline(),registration['command']):
                        raise RuntimeError('Registered parent identity mismatch')
                    parent=capture(process)
                    entry={**registration,'parent':parent,'children':{},'action':None};entries[key]=entry
                except (psutil.NoSuchProcess,psutil.AccessDenied):continue
            if entry.get('action'):continue
            parent=exact(entry['parent'])
            if parent is not None:
                try:
                    for child in parent.children(recursive=True):
                        try:
                            command=child.cmdline()
                            if child.create_time()>=entry['parent']['create_time']-.01 and compatible(command,entry['command']):
                                entry['children'][str(child.pid)]=capture(child)
                        except (psutil.NoSuchProcess,psutil.AccessDenied):pass
                except psutil.NoSuchProcess:pass
            reason=cancellations.get(key)
            if reason or time.time()>entry['deadline']:
                entry['action']='OBSOLETE_ROUTE_CANCELLED' if reason else 'HARD_WALL_TIMEOUT'
                entry['reason']=reason or 'External wall deadline reached';entry['stopped_at_epoch']=time.time()
                entry['stopped_pids']=[]
                # Actual solver descendants first; the venv redirector last.
                for record in list(entry['children'].values())+[entry['parent']]:
                    process=exact(record)
                    if process is not None:
                        try:
                            process.terminate();entry['stopped_pids'].append(process.pid)
                            try:process.wait(timeout=2)
                            except psutil.TimeoutExpired:
                                again=exact(record)
                                if again is not None:again.kill();again.wait(timeout=2)
                        except psutil.NoSuchProcess:pass
                folder=Path(entry['case_folder'])
                entry['preserved_files']={p.relative_to(folder).as_posix():hashlib.sha256(p.read_bytes()).hexdigest()
                    for p in sorted(folder.rglob('*')) if p.is_file()}
                print(key,entry['action'],entry['stopped_pids'],flush=True)
        ledger['last_epoch']=time.time();atomic(output/'GUARD_LEDGER.json',ledger)
        time.sleep(1)
    atomic(output/'GUARD_LEDGER.json',ledger)

if __name__=='__main__':main()
