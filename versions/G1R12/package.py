from pathlib import Path
import json,hashlib,shutil,csv,zipfile,subprocess
ROOT=Path(__file__).resolve().parent
REPO=Path('C:/Users/link_/Downloads/excel-vba-rpfem_G1R11_20261008')
def digest(p):return hashlib.sha256(p.read_bytes()).hexdigest()
before=json.loads((ROOT/'source_identity.json').read_text(encoding='utf-8'))
assert digest(Path(before['path']))==before['sha256']
oldzip=REPO/'downloads/RPFEM_G1R11_Staged_20261007.zip'
assert digest(oldzip)=='f1e2b475af8f6496de104ac5083ba8cafcd079789e67931300f668361f6f2b0c'
protected={'G1R11_xlsm_sha256':digest(Path(before['path'])),'G1R11_zip_sha256':digest(oldzip),'originals_unchanged':True}
(ROOT/'results/protected_originals.json').write_text(json.dumps(protected,indent=2),encoding='utf-8')
checkroot=Path('C:/Users/link_/Downloads/RPFEM_Codex_P2S_0945_Task')
out=subprocess.check_output([str(Path('C:/Users/link_/Downloads/RPFEM_C0_UpperMesh_Codex_Plan/.venv/Scripts/python.exe')),'-X','utf8','tools/verify_baseline.py'],cwd=checkroot,text=True,encoding='utf-8')
check=json.loads(out);assert check['baseline_intact'] and check['checked_files']==32
(ROOT/'results/baseline_after.json').write_text(json.dumps(check,indent=2),encoding='utf-8')
with (ROOT/'ISSUE_FIX_MAP.csv').open('w',encoding='utf-8-sig',newline='') as f:
    w=csv.writer(f);w.writerow(['issue','verdict','modules','fix','evidence'])
    w.writerow([1,'妥当','RPX_Control;RPX_Output','メッシュ表と図の前に上界witnessを復元;厳密所有者検査;下界は別所有者','results/native_G1R11.json;results/native_G1R12.json;results/native_bearing.json'])
    w.writerow([2,'妥当','RPX_Robust','pivot付き実エラー形式のみ追加;詳細保持;既存救済・停止条件維持','results/native_G1R11.json;results/native_G1R12.json;results/independent_audits.json'])
rows=[
 ('PY_PRETEST','Python simulation/static','PASS','3 assertions; 28 grammar vectors','results/python_pretest.json'),
 ('NATIVE_OLD_REPRO','Excel synthetic owner test','PASS','旧版のError9・同数座標不一致・救済0回を再現','results/native_G1R11.json'),
 ('NATIVE_OUTPUT','Excel synthetic owner test','PASS','4 owner variants','results/native_G1R12.json'),
 ('NATIVE_OUTPUT_GUARDS','Excel failure injection','PASS','7 owner/model/context guards','results/native_G1R12.json'),
 ('NATIVE_TOKENS','Excel grammar test','PASS','28 vectors + cancel/audit flags','results/native_G1R12.json'),
 ('NATIVE_FACTOR_RESCUE','Excel real small solver + FactorSystem failure injection','PASS','HSD_RESCUE exactly once; OPTIMAL; audit','results/native_G1R12.json'),
 ('NATIVE_STOP_GUARDS','Excel failure injection','PASS','memory/subscript/cancel/unknown/audit/model/HSD memory propagate','results/native_G1R12.json'),
 ('NATIVE_ROOTS','Excel physical root regression','PASS','C5 phi=psi45; Fs/fields/mesh/calls identical','results/native_G1R11.json;results/native_G1R12.json'),
 ('PY_NATIVE_AUDIT','Python independent physical audit','PASS','7 native fields; lower 50digit audits','results/independent_audits.json'),
 ('NATIVE_WORKFLOW','Excel actual adaptive RPX_Run','PASS','small C5; equal bounds/table; GAP target still unmet','results/native_workflow.json'),
 ('NATIVE_BEARING','Excel actual bearing workflow','PASS','nonadaptive/adaptive small bearing multiplier3.5/pressure2','results/native_bearing.json'),
 ('NATIVE_PRODUCTION','Excel compile/save/reopen/source check','PASS','46 modules; no test probes; inputs unchanged','results/production_workbook.json'),
 ('PROTECTION','SHA256/static','PASS','G1R11 artifacts + 32 protected P2S files unchanged','results/protected_originals.json;results/baseline_after.json'),
 ('P06N_FULL','Excel full adaptive','NOT_RUN','全P06N・GAP1%・時間短縮率未確認','TEST_PROCEDURES.md'),
 ('FULL25','Excel full regression','NOT_RUN','25 cases excluding pore pressure; no inferred PASS','excel_acceptance_g1r12.csv'),
 ('RIGID_FULL','Excel full regression','NOT_RUN','剛体・支持力全回帰未実施','TEST_PROCEDURES.md')]
with (ROOT/'TEST_RESULTS.csv').open('w',encoding='utf-8-sig',newline='') as f:
    w=csv.writer(f);w.writerow(['test','scope','status','description','evidence']);w.writerows(rows)
old=(REPO/'versions/G1R11/excel_acceptance_g1r11.csv').read_text(encoding='utf-8')
assert old.count('NOT_RUN')==25
(ROOT/'excel_acceptance_g1r12.csv').write_text(old,encoding='utf-8',newline='\n')

selected=[]
for folder in ['src','original','import_cp932','rollback_cp932','validation_runtime']:
    selected.extend(p for p in (ROOT/folder).rglob('*') if p.is_file() and '__pycache__' not in p.parts and p.suffix!='.pyc')
selected.extend(p for p in (ROOT/'tests').iterdir() if p.is_file() and p.suffix in ['.py','.bas','.txt','.md','.json'])
selected=[p for p in selected if p.name!='compile_failed_source.txt']
keep=['python_pretest.json','native_G1R11.json','native_G1R12.json','independent_audits.json','native_workflow.json','native_bearing.json','production_workbook.json','baseline_before.json','baseline_after.json','protected_originals.json']
selected.extend(ROOT/'results'/n for n in keep)
selected.extend((ROOT/'results').glob('export_*.bas'))
earliest=None
for version in ['G1R11','G1R12']:
    p=ROOT/'results'/f'native_{version}.json';rec=json.loads(p.read_text(encoding='utf-8'));assert rec['pass']
    exports=Path(rec['exports']);selected.extend(x for x in exports.iterdir() if x.is_file())
    rec['exports']=exports.relative_to(ROOT).as_posix()
    p.write_text(json.dumps(rec,ensure_ascii=False,indent=2),encoding='utf-8')
    stamp='20261009_'+exports.name.split('_')[1]
    earliest=stamp if earliest is None else min(earliest,stamp)
selected.extend(p for p in (ROOT/'tests/RPFEM_logs').iterdir() if p.is_file() and p.name[:15]>=earliest)
selected.extend(p for p in ROOT.iterdir() if p.is_file() and p.suffix in ['.xlsm','.md','.csv','.json','.patch','.txt','.py'])
selected=sorted(set(selected))
manifest=[{'path':p.relative_to(ROOT).as_posix(),'bytes':p.stat().st_size,'sha256':digest(p)} for p in selected]
(ROOT/'artifact_manifest.json').write_text(json.dumps({'files':manifest,'full_P06N':'NOT_RUN','full25':'NOT_RUN'},ensure_ascii=False,indent=2),encoding='utf-8')
selected.append(ROOT/'artifact_manifest.json')
target=REPO/'versions/G1R12';assert not target.exists();target.mkdir(parents=True)
for p in selected:
    dest=target/p.relative_to(ROOT);dest.parent.mkdir(parents=True,exist_ok=True);shutil.copy2(p,dest)
zipname='RPFEM_G1R12_IssueFix_20261009.zip';zp=REPO/'downloads'/zipname
with zipfile.ZipFile(zp,'w',zipfile.ZIP_DEFLATED,compresslevel=9) as z:
    for p in selected:z.write(p,p.relative_to(ROOT).as_posix())
with zipfile.ZipFile(zp) as z:
    assert z.testzip() is None and len(z.namelist())==len(selected)
    for row in manifest:assert hashlib.sha256(z.read(row['path'])).hexdigest()==row['sha256']
xl=next(p for p in selected if p.suffix=='.xlsm')
with (REPO/'downloads/SHA256SUMS.txt').open('a',encoding='utf-8',newline='\n') as f:
    f.write(digest(zp)+'  '+zipname+'\n'+digest(xl)+'  ../versions/G1R12/'+xl.name+'\n')
for p in [zp,xl]:
    dest=Path('C:/Users/link_/Desktop')/p.name;assert not dest.exists();shutil.copy2(p,dest)
print(json.dumps({'files':len(selected),'zip_bytes':zp.stat().st_size,'zip_sha256':digest(zp),'xlsm_sha256':digest(xl),'baseline32_unchanged':True,'full25':'NOT_RUN'},indent=2))
