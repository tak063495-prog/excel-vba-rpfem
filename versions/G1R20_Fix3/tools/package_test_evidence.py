"""Archive completed owned-test records; never replace an existing evidence ZIP."""
from pathlib import Path
import argparse,hashlib,json,zipfile

ROOT=Path(__file__).resolve().parents[1]
REPO=ROOT.parents[1]

def sha(raw):return hashlib.sha256(raw).hexdigest()

def main():
    p=argparse.ArgumentParser();p.add_argument('--output',type=Path,required=True)
    p.add_argument('--run',type=Path,action='append',required=True)
    a=p.parse_args();out=a.output.resolve()
    if out.exists():raise FileExistsError('Completed evidence must not be replaced')
    files={}
    summaries=[]
    for run in a.run:
        run=run.resolve()
        if not run.is_relative_to(ROOT/'results'):raise ValueError('Run must be inside this release results directory')
        record=run/'summary.json'
        if record.exists():
            summary=json.loads(record.read_text(encoding='utf8'))
            sessions = summary.get('sessions', summary.get('stages', []))
            exited = summary.get('owned_excel_exited') or (bool(sessions) and all(s.get('owned_excel_exited') for s in sessions))
            later_exit = None
            if not exited:
                # Preserve an original teardown failure. A separate, dated
                # observation may establish that the exact owner later exited;
                # check it again now and never terminate a different process.
                import psutil
                later_path = run/'POST_RUN_EXIT_CHECK.json'
                if not later_path.exists():raise ValueError('Owned Excel has not exited: '+str(run))
                later_exit = json.loads(later_path.read_text(encoding='utf8'))
                owner = summary.get('excel_owner', {})
                pid = summary.get('owned_excel_pid', owner.get('pid'))
                birth = summary.get('owned_excel_creation_time', owner.get('birth', owner.get('create_time')))
                assert pid and birth and later_exit.get('pid', later_exit.get('owned_pid'))==pid
                assert later_exit.get('creation_time', later_exit.get('owned_birth'))==birth
                assert later_exit.get('original_owner_exited', later_exit.get('ownership_exit_verified')) and later_exit['observed_utc']
                try:
                    process = psutil.Process(pid)
                    if process.create_time()==birth:raise ValueError('Original Excel owner is still alive: '+str(run))
                except psutil.NoSuchProcess:pass
            summaries.append({'path':record.relative_to(REPO).as_posix(),
                              'status':summary.get('status', summary.get('state')),'sha256':sha(record.read_bytes()),
                              'original_owned_excel_exited':bool(exited),'later_exit_observation':later_exit})
        else:
            record=run/'delivery/WORKBOOK_VALIDATION.json'
            summary=json.loads(record.read_text(encoding='utf8'))
            if summary['state']!='PASS' or not summary['owned_excel_exited']:raise ValueError('Book build incomplete')
            summaries.append({'path':record.relative_to(REPO).as_posix(),
                              'status':summary['state'],'sha256':sha(record.read_bytes())})
        for path in sorted(run.rglob('*')):
            if not path.is_file() or '__pycache__' in path.parts or path.name.startswith('~$'):continue
            files[path.relative_to(REPO).as_posix()]=path.read_bytes()
    paths={name:{'bytes':len(raw),'sha256':sha(raw)} for name,raw in files.items()}
    manifest=json.dumps({'schema':'rpfem.g20fix3.execution.v1','files':paths,'summaries':summaries},indent=2).encode('utf8')
    files['MANIFEST.json']=manifest
    out.parent.mkdir(parents=True,exist_ok=True)
    with zipfile.ZipFile(out,'x',zipfile.ZIP_DEFLATED,compresslevel=6) as z:
        for name,raw in sorted(files.items()):z.writestr(name,raw)
    with zipfile.ZipFile(out) as z:
        if z.testzip() is not None:raise ValueError('Evidence CRC failed')
        for name,raw in files.items():
            if sha(z.read(name))!=sha(raw):raise ValueError('Evidence bytes differ: '+name)
    result={'status':'PASS','archive':out.name,'bytes':out.stat().st_size,'sha256':sha(out.read_bytes()),
            'members':len(files),'source_file_count':len(paths),'member_sha256_and_crc':'PASS',
            'run_summaries':summaries,'scope':'Owned-test source snapshots, configured copies, raw fields and logs; no original G1R20 evidence replaced'}
    out.with_suffix('.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf8')
    print(json.dumps(result))

if __name__=='__main__':main()
