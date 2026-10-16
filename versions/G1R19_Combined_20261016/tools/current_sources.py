"""Compose an unpublished ten-module correction over the immutable G1R19."""
from pathlib import Path
import argparse,hashlib,json

ROOT=Path(__file__).resolve().parents[1];BASE=ROOT.parent/'G1R19'
UPSTREAM=[ROOT.parent/'G1R19_2pct',ROOT.parent/'G1R19_WaterCommon_20261016',ROOT.parent/'G1R19_StateReset_20261016']
def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def materialize_overlay():
    if (ROOT/'SOURCE_MANIFEST.json').exists():raise FileExistsError('Overlay already exists')
    files={};parents={}
    for version in UPSTREAM:
        parents[version.name]=sha(version/'SOURCE_MANIFEST.json')
        for p in (version/'src/vba').glob('*.bas'):files[p.name]=(version,p)
    assert len(files)==10
    records={}
    for name,(version,p) in files.items():
        imp=version/'import'/name
        assert p.read_text(encoding='utf8')==imp.read_text(encoding='cp932')
        for sub,source in [('src/vba',p),('import',imp)]:
            dest=ROOT/sub/name;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(source.read_bytes())
        records[name]={'source':str(p.relative_to(ROOT.parent)),'utf8_sha256':sha(p),'cp932_sha256':sha(imp)}
    data=dict(schema='rpfem.combined.overlay.v1',candidate_only=True,release=False,
        scope='Four accuracy corrections, pinned prescribed-water common mesh, four explicit entry calls',
        p06_target_claimed=False,upstream_manifest_sha256=parents,modules=records,
        warning='D08 cold proof used the seven-module composition before the four explicit-call changes; do not relabel it as a ten-module numerical run')
    (ROOT/'SOURCE_MANIFEST.json').write_text(json.dumps(data,indent=2)+'\n',encoding='utf8')

def module_paths():
    manifest=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    files={p.stem:p for p in (BASE/'src/vba').glob('*.bas')}
    files.update({p.stem:p for p in (ROOT/'src/vba').glob('*.bas')})
    assert len(files)==62 and len(manifest['modules'])==10
    for name,p in files.items():
        version=ROOT if p.is_relative_to(ROOT) else BASE
        assert p.read_text(encoding='utf8')==(version/'import'/p.name).read_text(encoding='cp932')
        if version==ROOT:
            row=manifest['modules'][p.name]
            assert sha(p)==row['utf8_sha256'] and sha(ROOT/'import'/p.name)==row['cp932_sha256']
    return files

def generate(output):
    output=Path(output).resolve()
    if output.exists():raise FileExistsError('Refusing to replace candidate: '+str(output))
    records={}
    for name,p in module_paths().items():
        version=ROOT if p.is_relative_to(ROOT) else BASE
        for sub,source in [('src/vba',p),('import',version/'import'/p.name)]:
            dest=output/sub/p.name;dest.parent.mkdir(parents=True,exist_ok=True);dest.write_bytes(source.read_bytes())
        records[p.name]={'utf8_sha256':sha(output/'src/vba'/p.name),'cp932_sha256':sha(output/'import'/p.name),'source':str(p.relative_to(ROOT.parent))}
    data=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    data['overlay_manifest_sha256']=sha(ROOT/'SOURCE_MANIFEST.json');data['modules']=records
    (output/'SOURCE_MANIFEST.json').write_text(json.dumps(data,indent=2)+'\n',encoding='utf8')
    return dict(status='PASS',output=str(output),modules=len(records))

if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--materialize-overlay',action='store_true');p.add_argument('--output-root',type=Path)
    args=p.parse_args()
    if args.materialize_overlay:materialize_overlay()
    if args.output_root:print(json.dumps(generate(args.output_root)))
