"""Prepare exact Git objects for the connected GitHub API; no network writes."""
from pathlib import Path
import base64,hashlib,json,subprocess,sys
from prepare_inputs import ROOT,REPO
D=ROOT/'github_transfer'
def git(*args):return subprocess.check_output(['git',*args],cwd=REPO)
def main():
    mode=sys.argv[1]
    if mode=='prepare':
        D.mkdir(exist_ok=True)
        summary=json.loads((ROOT/'results/final_summary.json').read_text(encoding='utf-8'))
        assert summary['all25_native_executed']
        old={line.split()[2] for line in git('ls-tree','-r','HEAD').decode().splitlines()}
        changed=[p.decode() for p in git('diff','--cached','--name-only','-z').split(b'\0') if p]
        assert changed
        assert all(p in ['README.md','.gitattributes','docs/USAGE_JA.md','docs/NOTES_JA.md','downloads/SHA256SUMS.txt'] or
                   p.startswith('validation/G1R12_native25_20261009/') or
                   p.startswith('downloads/RPFEM_G1R12_Native25_') for p in changed)
        entries=[];needed={}
        for path in changed:
            digest=git('rev-parse',':'+path).decode().strip();raw=git('cat-file','blob',digest)
            assert hashlib.sha1(b'blob '+str(len(raw)).encode()+b'\0'+raw).hexdigest()==digest
            item={'path':path,'mode':'100644','type':'blob','sha':digest}
            if digest not in old:
                try:
                    assert len(raw)<=32768 and Path(path).suffix not in ['.bin','.npz','.xlsm','.zip','.png']
                    content=raw.decode('utf-8');assert content.encode('utf-8')==raw
                    item.pop('sha');item['content']=content
                except (AssertionError,UnicodeDecodeError):
                    (D/(digest+'.bin')).write_bytes(raw);needed[digest]={'sha':digest,'bytes':len(raw)}
            entries.append(item)
        meta={'base_commit':git('rev-parse','HEAD').decode().strip(),
              'base_tree':git('rev-parse','HEAD^{tree}').decode().strip(),
              'expected_tree_sha':git('write-tree').decode().strip(),
              'entries':entries,'needed_blobs':list(needed.values())}
        (D/'manifest.json').write_text(json.dumps(meta,ensure_ascii=False),encoding='utf-8')
        print(json.dumps({k:v for k,v in meta.items() if k!='entries'}))
    elif mode=='meta':
        m=json.loads((D/'manifest.json').read_text(encoding='utf-8'));m.pop('entries');print(json.dumps(m))
    elif mode=='entries':
        m=json.loads((D/'manifest.json').read_text(encoding='utf-8'));start,n=int(sys.argv[2]),int(sys.argv[3])
        print(json.dumps(m['entries'][start:start+n],ensure_ascii=False))
    elif mode=='chunk':
        digest=sys.argv[2];assert len(digest)==40 and all(c in '0123456789abcdef' for c in digest)
        offset,size=int(sys.argv[3]),int(sys.argv[4]);assert offset>=0 and 0<size<=294912
        with (D/(digest+'.bin')).open('rb') as f:f.seek(offset);raw=f.read(size)
        print(json.dumps({'bytes':len(raw),'data':base64.b64encode(raw).decode()}))
    else:raise ValueError(mode)

if __name__=='__main__':main()
