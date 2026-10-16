"""The complete G1R20 src/vba is authoritative; no scratch overlay dependency."""
from pathlib import Path
import hashlib,json
ROOT=Path(__file__).resolve().parents[1]
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def generate(output):
    output=Path(output).resolve()
    if output.exists():raise FileExistsError('Refusing to replace candidate output')
    manifest=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    assert len(manifest['modules'])==63
    for name,row in manifest['modules'].items():
        for sub,key in [('src/vba','utf8_sha256'),('import','cp932_sha256')]:
            p=ROOT/sub/name;assert sha(p)==row[key],name
            dest=output/sub/name;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(p.read_bytes())
    (output/'SOURCE_MANIFEST.json').write_bytes((ROOT/'SOURCE_MANIFEST.json').read_bytes())
    return {'status':'PASS','modules':63,'output':str(output)}
