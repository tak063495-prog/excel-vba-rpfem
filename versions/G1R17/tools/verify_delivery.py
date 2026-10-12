"""Read-only source/workbook/protected-file verification and unified diffs."""
import json,difflib
from support import ROOT,REPO,BASE,SOURCE,write,sha
from inspect_saved import macros,canonical,saved_cells
def main():
 old=macros(SOURCE);source={p.stem:p.read_text(encoding='utf8')for p in (ROOT/'src/vba').iterdir()if p.suffix in ['.bas','.cls']}
 assert len(old)==70 and len(source)==71 and set(source)-set(old)=={'RPX_NativeMesh'}
 changes=[n for n in old if canonical(old[n])!=canonical(source[n])];assert set(changes)=={'RPX_Mesh','RPX_Native'}
 protected=json.loads((ROOT/'results/baseline.json').read_text(encoding='utf8'))['protected'];assert all(sha(REPO/p)==h for p,h in protected.items())
 records=[]
 for path in sorted((ROOT/'delivery').glob('*.xlsm')):
  new=macros(path);assert set(new)==set(source)
  assert all(canonical(new[k])==canonical(v)for k,v in source.items()),path
  assert all(canonical(old[k])==canonical(new[k])for k in old if k.startswith('Sheet')or k=='ThisWorkbook')
  cells=saved_cells(path);settings=cells['設定'];assert settings['B5']==200 and settings['B7']==0 and cells['浸透入力']['D4']==0
  result=cells['解析結果'];assert 0<result['B4']<=result['C4']
  records.append({'file':path.name,'sha256':sha(path),'modules':len(new),'source_match':True,'document_modules_unchanged':True,'lower':result['B4'],'upper':result['C4'],'settings_coarse':True})
 assert len(records)==2
 for name in ['RPX_Mesh','RPX_Native','RPX_NativeMesh']:
  data=(ROOT/'import'/(name+'.bas')).read_bytes();assert data.decode('cp932').replace('\r\n','\n')==source[name];assert b'\r\n'in data
 diff=[]
 for name in sorted(changes+['RPX_NativeMesh']):
  p=BASE/'src/vba'/(name+'.bas');a=p.read_text(encoding='utf8').splitlines(True)if p.exists()else []
  diff.extend(difflib.unified_diff(a,source[name].splitlines(True),fromfile='G1R16_DLL5/'+name+'.bas'if a else'/dev/null',tofile='G1R17/'+name+'.bas'))
 (ROOT/'changes.diff').write_text(''.join(diff),encoding='utf8',newline='\n')
 diff=[];changed_native=[]
 for p in sorted((ROOT/'native').glob('*')):
  if p.suffix not in ['.c','.cpp','.h','.asm','.cmd']:continue
  oldp=BASE/'native'/p.name;a=oldp.read_text(encoding='utf8').splitlines(True)if oldp.exists()else [];b=p.read_text(encoding='utf8').splitlines(True)
  if a!=b:changed_native.append(p.name)
  diff.extend(difflib.unified_diff(a,b,fromfile='G1R16_DLL5/native/'+p.name if a else'/dev/null',tofile='G1R17/native/'+p.name))
 (ROOT/'native_changes.diff').write_text(''.join(diff),encoding='utf8',newline='\n')
 assert set(changed_native)=={'build.cmd','rpfem_native.c','vba_fp.asm','rpfem_meshfan.cpp'}
 record={'PASS':True,'protected_originals':len(protected),'protected_hashes_match':True,'old_modules':len(old),'new_modules':len(source),'changed_existing_modules':changes,'added_modules':['RPX_NativeMesh'],'unchanged_existing_modules':len(old)-len(changes),'changed_native_files':changed_native,'DLL_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll'),'old_DLL_sha256':sha(BASE/'delivery/rpfem_native_x64.dll'),'CP932_CRLF_lossless':True,'workbooks':records}
 write(ROOT/'results/delivery_verification.json',record);print('DELIVERY VERIFICATION PASS',len(records),'books,',len(protected),'protected files',flush=True)
if __name__=='__main__':main()
