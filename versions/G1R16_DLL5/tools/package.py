"""Package verified final payload with byte-preserving hashes and desktop copy."""
import json,shutil,subprocess,sys,zipfile,hashlib,time
from pathlib import Path
from prepare import ROOT,REPO,write,sha

EXCLUDED_RESULTS={'fixed_fem_dll4_initial','fixed_fem_dll4_refine','fixed_fem_hlinear','levee_dll4_initial','fixed_q8_final','fixed_q8_fused','p06n_upper_profile'}
def selected(p):
 rel=p.relative_to(ROOT);parts=rel.parts
 if '__pycache__' in parts or p.name.startswith('~$'):return False
 if parts[0]=='delivery':return p.name in ['rpfem_native_x64.dll','RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm']
 if parts[0]=='native':return p.suffix in ['.c','.cpp','.h','.asm','.cmd','.txt']
 if parts[0]=='tests':return p.suffix in ['.bas','.txt','.cpp','.cmd']
 if parts[0]=='results':
  if p.suffix in ['.xlsm','.dll','.pyc'] or p.name.startswith(('working_before_update_','compile_')):return False
  if len(parts)>1 and parts[1] in EXCLUDED_RESULTS or 'nttrace' in p.name:return False
  if p.name in ['package_verification.json','publication_verification.json','desktop_delivery.json']:return False
  return p.suffix in ['.json','.bin','.tsv','.log','.txt','.csv']
 if parts[0]=='tools':return p.suffix in ['.py','.md'] and not p.name.startswith('implement_') and p.name not in ['add_solver_asm.py','stop_trial.py','probe_nt_trace.py']
 return p.suffix in ['.md','.diff','.bas','.cls','.txt','.json','.csv'] or p.name in ['.gitignore','.gitattributes']

def main():
 verification=json.loads((ROOT/'results/delivery_verification.json').read_text(encoding='utf-8'))
 assert verification['reopen_whole_project_compile']=='PASS' and verification['xlsm_sha256']==sha(ROOT/'delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm') and verification['dll_sha256']==sha(ROOT/'delivery/rpfem_native_x64.dll')
 assert json.loads((ROOT/'results/pipeline.json').read_text(encoding='utf-8'))['state']=='completed'
 core=json.loads((ROOT/'results/core_pipeline.json').read_text(encoding='utf-8'))
 assert core['state']=='completed' and len(core['runs'])==22 and all(r['exit_code']==0 for r in core['runs'])
 assert json.loads((ROOT/'results/fixed_comparison.json').read_text(encoding='utf-8'))['passed']
 assert json.loads((ROOT/'results/root_regressions.json').read_text(encoding='utf-8'))['passed_comparisons']==4
 protected=json.loads((ROOT/'results/baseline.json').read_text(encoding='utf-8'))['protected'];assert all(sha(REPO/p)==h for p,h in protected.items())
 oldtask=Path('C:/Users/link_/Downloads/RPFEM_Codex_P2S_0945_Task')
 r=subprocess.run([sys.executable,'tools/verify_baseline.py'],cwd=oldtask,capture_output=True,text=True,encoding='utf-8',check=True)
 historical=json.loads(r.stdout);assert historical['baseline_intact'];write(ROOT/'results/historical_baseline_preserved.json',historical)
 # Existing timings are immutable references, with their provenance copied to the ZIP.
 references=ROOT/'results/reference';references.mkdir(exist_ok=True)
 for name in ['vba_analysis.tsv','metadata.json']:
  source=REPO/'versions/G1R16_DLL_3/results/reference'/name;shutil.copyfile(source,references/name)
 files=sorted(p for p in ROOT.rglob('*') if p.is_file() and selected(p) and p.name not in ['MANIFEST.json','SHA256SUMS.txt'])
 manifest={str(p.relative_to(ROOT)).replace('\\','/'):sha(p) for p in files};write(ROOT/'MANIFEST.json',manifest)
 (ROOT/'SHA256SUMS.txt').write_text(''.join(h+'  '+p+'\n' for p,h in manifest.items()),encoding='utf-8',newline='\n');files.extend([ROOT/'MANIFEST.json',ROOT/'SHA256SUMS.txt'])
 dst=REPO/'downloads/RPFEM_G1R16_DLL5_20261012.zip';assert not dst.exists(),'Preserve prior package; choose a new release name instead.'
 with zipfile.ZipFile(dst,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=6) as z:
  for p in files:z.write(p,str(p.relative_to(ROOT)).replace('\\','/'))
 with zipfile.ZipFile(dst) as z:
  assert z.testzip() is None and len(z.namelist())==len(set(z.namelist()))==len(files)
  assert all(hashlib.sha256(z.read(p)).hexdigest()==h for p,h in manifest.items())
  assert sum(n.endswith('.xlsm') for n in z.namelist())==1 and sum(n.endswith('.dll') for n in z.namelist())==1
 artifacts=[('downloads/RPFEM_G1R16_DLL5_20261012.zip','RPFEM_G1R16_DLL5_20261012.zip'),('versions/G1R16_DLL5/delivery/rpfem_native_x64.dll','../versions/G1R16_DLL5/delivery/rpfem_native_x64.dll'),('versions/G1R16_DLL5/delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm','../versions/G1R16_DLL5/delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm')]
 sums=REPO/'downloads/SHA256SUMS.txt';before=sums.read_bytes();assert all(name.encode() not in before for _,name in artifacts)
 with sums.open('ab') as f:f.write(''.join(sha(REPO/p)+'  '+name+'\n' for p,name in artifacts).encode('utf-8'))
 assert sums.read_bytes().startswith(before)
 write(ROOT/'results/package_verification.json',{'zip_path':str(dst.relative_to(REPO)),'zip_sha256':sha(dst),'zip_bytes':dst.stat().st_size,'entries':len(files),'every_payload_hash_verified':True,'only_one_delivery_xlsm_and_dll':True,'historical32_and_protected1085_intact':True})
 local=Path('C:/Users/link_/Desktop/RPFEM_G1R16_DLL5_20261012_Final')
 if local.exists():local=local.with_name(local.name+'_'+str(time.time_ns()))
 local.mkdir()
 for p in [ROOT/'delivery/rpfem_native_x64.dll',ROOT/'delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm',ROOT/'README.md',ROOT/'REPORT.md',ROOT/'IMPORT_ROLLBACK_JA.md',ROOT/'IMPLEMENTATION_JA.md',dst]:shutil.copyfile(p,local/p.name)
 # The desktop copy is flat; its checksums must name the files actually present.
 local_files=sorted(p for p in local.iterdir() if p.is_file())
 (local/'SHA256SUMS.txt').write_text(''.join(sha(p)+'  '+p.name+'\n' for p in local_files),encoding='utf-8',newline='\n')
 write(ROOT/'results/desktop_delivery.json',{'folder':str(local),'files':{p.name:sha(p) for p in local.iterdir() if p.is_file()}})
 print('ZIP verified',dst.stat().st_size,'bytes;',sha(dst),'Desktop:',local,flush=True)
if __name__=='__main__':main()
