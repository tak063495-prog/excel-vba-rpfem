"""Verify/rebuild source byte metadata from the authoritative G1R19 modules."""
from pathlib import Path
import argparse,hashlib,json
ROOT=Path(__file__).resolve().parents[1]
def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()
def main():
    parser=argparse.ArgumentParser();parser.add_argument('--refresh-manifest',action='store_true');args=parser.parse_args()
    path=ROOT/'SOURCE_MANIFEST.json';manifest=json.loads(path.read_text(encoding='utf8'))
    modules={p.name:p for p in (ROOT/'src/vba').glob('*.bas')};assert len(modules)==62
    assert manifest['release']=='G1R19' and set(modules)==set(manifest['modules'])
    for name,source in modules.items():
        imported=ROOT/'import'/name;assert source.read_text(encoding='utf8')==imported.read_text(encoding='cp932'),name
        hashes={'utf8_sha256':sha(source),'cp932_sha256':sha(imported)}
        if args.refresh_manifest:manifest['modules'][name].update(hashes)
        else:assert all(manifest['modules'][name][key]==value for key,value in hashes.items()),name
    if args.refresh_manifest:
        manifest.update(release='G1R19',acceptance_gap=.02,unified_dll=True,direct_dll_default=0,direct_dll_preset=1,
                        changes=['Original LP before legacy solver; strict original primal/dual gates; bounded fresh original dual helper',
                                 'One DLL retaining both ABIs','Same-run exact-context terminal q8 pair reuse with fresh physical audits',
                                 '2-percent acceptance and speed preset; precision-from-start remains zero',
                                 'Issue12 quadrature/epoch cache guards and Issue15 hydrology stale-result fix retained'])
        for record in manifest['modules'].values():
            if record.get('overlay')=='G1R18_Speed2pct':record['overlay']='G1R19'
        path.write_text(json.dumps(manifest,indent=2)+'\n',encoding='utf8')
    print(json.dumps({'status':'PASS','modules':62,'source_manifest_sha256':sha(path)}))
if __name__=='__main__':main()
