"""G1R19 plus the four authoritative 2-percent overlay modules."""
from pathlib import Path
import argparse, hashlib, json

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT.parent/'G1R19'

def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def module_paths():
    files = {p.stem:p for p in (BASE/'src/vba').glob('*.bas')}
    files.update({p.stem:p for p in (ROOT/'src/vba').glob('*.bas')})
    if len(files) != 62:
        raise ValueError('Expected the complete 62-module program')
    for name,path in files.items():
        version = ROOT if path.is_relative_to(ROOT) else BASE
        if path.read_text(encoding='utf8') != (version/'import'/path.name).read_text(encoding='cp932'):
            raise ValueError('CP932 import mismatch: '+name)
    return files

def generate(output):
    output=Path(output).resolve()
    if output.exists():
        raise FileExistsError('Refusing to replace an existing candidate: '+str(output))
    records={}
    for name,path in module_paths().items():
        text=path.read_text(encoding='utf8').replace('\n','\r\n')
        for directory,encoding in [('src/vba','utf8'),('import','cp932')]:
            dest=output/directory/path.name;dest.parent.mkdir(parents=True,exist_ok=True)
            dest.write_bytes(text.encode(encoding))
        records[path.name]=dict(utf8_sha256=sha(output/'src/vba'/path.name),
                               cp932_sha256=sha(output/'import'/path.name),
                               source=str(path.relative_to(ROOT.parent)))
    manifest=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    manifest['modules']=records
    (output/'SOURCE_MANIFEST.json').write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf8')
    return dict(status='PASS',output=str(output),modules=len(records))

if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--output-root',required=True,type=Path)
    print(json.dumps(generate(parser.parse_args().output_root)))
