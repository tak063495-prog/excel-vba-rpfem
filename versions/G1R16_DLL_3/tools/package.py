import json,shutil,subprocess,sys,zipfile,hashlib,time
from pathlib import Path
from prepare import ROOT,REPO,write,sha
def selected(p):
    rel=p.relative_to(ROOT);parts=rel.parts
    if '__pycache__' in parts or p.name.startswith('~$'):return False
    if parts[0]=='delivery':return p.name in ['rpfem_native_x64.dll','RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm']
    if parts[0]=='native':return p.suffix in ['.c','.cpp','.h','.asm','.cmd','.txt']
    if parts[0]=='tests':return p.suffix in ['.bas','.txt','.cpp','.cmd']
    if parts[0]=='results':
        if p.suffix in ['.xlsm','.dll','.pyc']:return False
        if p.name.startswith(('working_before_update_','compile_')):return False
        if len(parts)>1 and parts[1] in ['fixed_fem','fixed_fem_verified','levee_native','levee_native_final','levee_native_checked']:return False
        if 'nttrace' in p.name:return False
        return p.suffix in ['.json','.bin','.tsv','.log','.txt','.csv']
    if parts[0]=='tools':return p.suffix=='.py' and p.name not in ['stop_trial.py','probe_nt_trace.py']
    return p.suffix in ['.md','.diff','.bas','.cls','.txt'] or p.name in ['.gitignore','.gitattributes']
def main():
    outputs={}
    for name in ['order','lu']:
        exe=ROOT/'tests'/f'{name}_failure_injection.exe'
        r=subprocess.run([str(exe)],capture_output=True,text=True,check=True)
        outputs[name]={'exit_code':r.returncode,'result':r.stdout.strip(),'source_sha256':sha(ROOT/'tests'/f'{name}_failure_injection.cpp')}
    write(ROOT/'results/failure_injection.json',outputs)
    protected=json.loads((ROOT/'results/baseline.json').read_text(encoding='utf-8'))['protected'];assert all(sha(REPO/p)==h for p,h in protected.items())
    original=Path('C:/Users/link_/Downloads/RPFEM_Codex_P2S_0945_Task')
    r=subprocess.run([sys.executable,'tools/verify_baseline.py'],cwd=original,capture_output=True,text=True,encoding='utf-8',check=True)
    historical=json.loads(r.stdout);assert historical['baseline_intact'];write(ROOT/'results/historical_baseline_preserved.json',historical)
    files=sorted(p for p in ROOT.rglob('*') if p.is_file() and selected(p) and p.name not in ['MANIFEST.json','SHA256SUMS.txt','package_verification.json'])
    manifest={str(p.relative_to(ROOT)).replace('\\','/'):sha(p) for p in files}
    write(ROOT/'MANIFEST.json',manifest)
    (ROOT/'SHA256SUMS.txt').write_text(''.join(h+'  '+p+'\n' for p,h in manifest.items()),encoding='utf-8')
    files.extend([ROOT/'MANIFEST.json',ROOT/'SHA256SUMS.txt'])
    dst=REPO/'downloads/RPFEM_G1R16_DLL_3_20261012.zip'
    if dst.exists():
        assert '--update' in sys.argv
        previous=json.loads((ROOT/'results/package_verification.json').read_text(encoding='utf-8'));assert sha(dst)==previous['zip_sha256']
        shutil.copyfile(dst,ROOT/'results'/('package_before_update_'+str(time.time_ns())+'.zip'))
    with zipfile.ZipFile(dst,'w',compression=zipfile.ZIP_DEFLATED,compresslevel=6) as z:
        for p in files:z.write(p,str(p.relative_to(ROOT)).replace('\\','/'))
    with zipfile.ZipFile(dst) as z:
        assert z.testzip() is None
        assert len(z.namelist())==len(set(z.namelist()))==len(files)
        assert all(hashlib.sha256(z.read(p)).hexdigest()==h for p,h in manifest.items())
        assert sum(n.endswith('.xlsm') for n in z.namelist())==1 and sum(n.endswith('.dll') for n in z.namelist())==1
    rels=['downloads/RPFEM_G1R16_DLL_3_20261012.zip','versions/G1R16_DLL_3/delivery/rpfem_native_x64.dll','versions/G1R16_DLL_3/delivery/RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm']
    sum_paths=[str(Path('..')/rel).replace('\\','/') if not rel.startswith('downloads/') else rel[len('downloads/'):] for rel in rels]
    sums=REPO/'downloads/SHA256SUMS.txt'
    if '--update' in sys.argv:
        lines=sums.read_text(encoding='utf-8').splitlines();sums.write_bytes(''.join(line+'\n' for line in lines if not any(line.endswith('  '+rel) for rel in rels+sum_paths)).encode('utf-8'))
    for rel,sum_path in zip(rels,sum_paths):
        with sums.open('a',encoding='utf-8',newline='\n') as f:f.write(sha(REPO/rel)+'  '+sum_path+'\n')
    write(ROOT/'results/package_verification.json',{'zip_path':str(dst.relative_to(REPO)),'zip_sha256':sha(dst),'zip_bytes':dst.stat().st_size,'entries':len(files),'every_payload_hash_verified':True,'only_one_delivery_xlsm_and_dll':True})
    # A separate desktop delivery; no modification of the user's existing books.
    local=Path('C:/Users/link_/Desktop/RPFEM_G1R16_DLL_3_20261012_Final' if '--update' in sys.argv else 'C:/Users/link_/Desktop/RPFEM_G1R16_DLL_3_20261012')
    if local.exists():local=local.with_name(local.name+'_'+str(time.time_ns()))
    local.mkdir()
    for p in [ROOT/'delivery/rpfem_native_x64.dll',ROOT/'delivery/RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm',ROOT/'README.md',ROOT/'REPORT.md',ROOT/'IMPORT_ROLLBACK_JA.md',ROOT/'IMPLEMENTATION_JA.md',ROOT/'SHA256SUMS.txt',dst]:shutil.copyfile(p,local/p.name)
    print('ZIP verified',dst.stat().st_size,'bytes',len(files),'files;',sha(dst),flush=True)
    print('Desktop delivery:',local,flush=True)
if __name__=='__main__':main()
