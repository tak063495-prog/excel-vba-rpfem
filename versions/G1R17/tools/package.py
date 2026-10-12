"""Package both actual result books, new DLL, complete source and evidence."""
import json,hashlib,subprocess,zipfile,shutil
from pathlib import Path
from support import ROOT,REPO,write,sha
ZIP='RPFEM_G1R17_20261012.zip'
def main():
 assert json.loads((ROOT/'results/acceptance.json').read_text(encoding='utf8'))['state']=='PASS'
 assert json.loads((ROOT/'results/reopen_delivery.json').read_text(encoding='utf8'))['PASS']
 assert json.loads((ROOT/'results/coarse_python_verification.json').read_text(encoding='utf8'))['PASS']
 allowed=[]
 for p in sorted(ROOT.rglob('*')):
  if not p.is_file()or '__pycache__'in p.parts or p.name in ['SHA256SUMS.txt','package_manifest.json','package_verification.json','publication_verification.json']:continue
  rel=p.relative_to(ROOT);suffix=p.suffix.lower()
  if suffix in ['.obj','.exp','.lib','.exe','.pyc']:continue
  if suffix in ['.xlsm','.dll']and rel.parts[0]!='delivery':continue
  if rel.parts[0]=='results'and p.name.startswith('nttrace'):continue
  allowed.append(p)
 books=[p for p in allowed if p.suffix=='.xlsm'];dlls=[p for p in allowed if p.suffix=='.dll'];assert len(books)==2 and len(dlls)==1
 manifest={str(p.relative_to(ROOT)).replace('\\','/'):{'sha256':sha(p),'bytes':p.stat().st_size}for p in allowed}
 write(ROOT/'package_manifest.json',{'version':'G1R17','ABI':'0x00170005','files':manifest})
 checksum=''.join(v['sha256']+'  '+k+'\n'for k,v in manifest.items());checksum+=sha(ROOT/'package_manifest.json')+'  package_manifest.json\n'
 (ROOT/'SHA256SUMS.txt').write_text(checksum,encoding='utf8')
 dest=REPO/'downloads'/ZIP
 with zipfile.ZipFile(dest,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=9)as z:
  for p in allowed+[ROOT/'package_manifest.json',ROOT/'SHA256SUMS.txt']:z.write(p,'G1R17/'+str(p.relative_to(ROOT)).replace('\\','/'))
 with zipfile.ZipFile(dest)as z:
  assert z.testzip()is None
  for key,v in manifest.items():assert hashlib.sha256(z.read('G1R17/'+key)).hexdigest()==v['sha256']
  assert len([n for n in z.namelist()if n.endswith('.xlsm')])==2 and len([n for n in z.namelist()if n.endswith('.dll')])==1
 write(ROOT/'results/package_verification.json',{'PASS':True,'zip':str(dest.relative_to(REPO)),'sha256':sha(dest),'bytes':dest.stat().st_size,'entries':len(allowed)+2,'workbooks':[p.name for p in books],'dlls':[p.name for p in dlls],'manifest_all_files_hashes_match':True})
 sums=REPO/'downloads/SHA256SUMS.txt';s=subprocess.check_output(['git','show','83debbecdc5f6376988aa1d93ebf1e6ecddf7477:downloads/SHA256SUMS.txt'],cwd=REPO)
 names=[dest,*books,*dlls]
 # Preserve previous release bytes and their original line endings exactly.
 newline=b'\r\n'if b'\r\n'in s else b'\n'
 if not s.endswith(b'\n'):s+=newline
 for p in names:
  name=p.name if p.parent==sums.parent else '../'+str(p.relative_to(REPO)).replace('\\','/');s+=(sha(p)+'  '+name).encode('utf8')+newline
 sums.write_bytes(s)
 desktop=Path('C:/Users/link_/Desktop/RPFEM_G1R17_20261012');desktop.mkdir(exist_ok=True)
 for p in names:shutil.copyfile(p,desktop/p.name);assert sha(desktop/p.name)==sha(p)
 for n in ['README.md','REPORT.md','IMPORT_ROLLBACK_JA.md','TEST_RESULTS.csv','SHA256SUMS.txt']:shutil.copyfile(ROOT/n,desktop/n)
 print('PACKAGE PASS',len(allowed)+2,'entries;',dest.stat().st_size,'bytes;SHA256',sha(dest),flush=True)
if __name__=='__main__':main()
