"""The complete G1R20 src/vba is authoritative; no scratch overlay dependency."""
from pathlib import Path
import hashlib,json
ROOT=Path(__file__).resolve().parents[1]
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def bundle_key(manifest):
    payload={name:manifest[name] for name in ('modules','workbook_modules')}
    return hashlib.sha256(json.dumps(payload,sort_keys=True,separators=(',',':')).encode('utf8')).hexdigest()

def generate(output):
    output=Path(output).resolve()
    if output.exists():raise FileExistsError('Refusing to replace candidate output')
    manifest=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    assert len(manifest['modules'])==63
    for name,row in manifest['modules'].items():
        for sub,key in [('src/vba','utf8_sha256'),('import','cp932_sha256')]:
            p=ROOT/sub/name;assert sha(p)==row[key],name
            dest=output/sub/name;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(p.read_bytes())
    assert set(manifest['workbook_modules'])=={'ThisWorkbook'}
    for row in manifest['workbook_modules'].values():
        for path_key,hash_key in [('utf8_path','utf8_sha256'),('cp932_path','cp932_sha256')]:
            p=ROOT/row[path_key];assert sha(p)==row[hash_key],row[path_key]
            dest=output/row[path_key];dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(p.read_bytes())
    assert bundle_key(manifest)==manifest['source_bundle_sha256']
    (output/'SOURCE_MANIFEST.json').write_bytes((ROOT/'SOURCE_MANIFEST.json').read_bytes())
    return {'status':'PASS','modules':63,'output':str(output)}
