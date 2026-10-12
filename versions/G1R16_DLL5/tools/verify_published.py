import argparse,hashlib,json,subprocess,urllib.request,time
from concurrent.futures import ThreadPoolExecutor
from prepare import ROOT,REPO,sha,write
def main():
 p=argparse.ArgumentParser();p.add_argument('commit');args=p.parse_args();assert len(args.commit)==40 and all(x in '0123456789abcdef' for x in args.commit)
 remote=subprocess.check_output(['git','ls-remote','origin','refs/heads/main'],cwd=REPO,text=True).split()[0];assert remote==args.commit
 paths=['downloads/RPFEM_G1R16_DLL5_20261012.zip','versions/G1R16_DLL5/delivery/rpfem_native_x64.dll','versions/G1R16_DLL5/delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm']
 def check(path):
  url='https://raw.githubusercontent.com/tak063495-prog/excel-vba-rpfem/'+args.commit+'/'+path
  for attempt in range(4):
   try:
    request=urllib.request.Request(url,headers={'User-Agent':'RPFEM-DLL5-release-check'})
    with urllib.request.urlopen(request,timeout=30) as r:
     assert r.status==200;h=hashlib.sha256();size=0
     while True:
      chunk=r.read(1024*1024)
      if not chunk:break
      h.update(chunk);size+=len(chunk)
    expected=sha(REPO/path);assert h.hexdigest()==expected and size==(REPO/path).stat().st_size
    return {'path':path,'url':url,'http_status':200,'bytes':size,'sha256':expected,'matches_local_artifact':True}
   except Exception:
    if attempt==3:raise
    time.sleep(2)
 with ThreadPoolExecutor(max_workers=3) as pool:records=list(pool.map(check,paths))
 write(ROOT/'results/publication_verification.json',{'main_commit':args.commit,'downloads':records,'pass':True});print('PUBLISHED 3 downloads HTTP200 + SHA256 PASS',args.commit,flush=True)
if __name__=='__main__':main()
