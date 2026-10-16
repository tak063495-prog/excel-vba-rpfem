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
            if not summary.get('owned_excel_exited'):raise ValueError('Owned Excel has not exited: '+str(run))
            summaries.append({'path':record.relative_to(REPO).as_posix(),
                              'status':summary.get('status'),'sha256':sha(record.read_bytes())})
        else:
            record=run/'delivery/WORKBOOK_VALIDATION.json'
            summary=json.loads(record.read_text(encoding='utf8'))
            if summary['state']!='PASS' or not summary['owned_excel_exited']:raise ValueError('Book build incomplete')
        for path in sorted(run.rglob('*')):
            if not path.is_file() or '__pycache__' in path.parts or path.name.startswith('~$'):continue
            files[path.relative_to(REPO).as_posix()]=path.read_bytes()
    paths={name:{'bytes':len(raw),'sha256':sha(raw)} for name,raw in files.items()}
    manifest=json.dumps({'schema':'rpfem.g20fix2.execution.v1','files':paths,'summaries':summaries},indent=2).encode('utf8')
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
