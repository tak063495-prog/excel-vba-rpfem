import csv,difflib,json,re,struct,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
from oletools.olevba import VBA_Parser
from prepare import ROOT,REPO,SOURCE,DLL3,sha,write
def canonical(text):
 text='\n'.join(s for s in text.splitlines() if not s.startswith('Attribute VB_'));result=[]
 for m in re.finditer(r'"(?:""|[^"])*"|\x27[^\n]*|[A-Za-z_][A-Za-z_0-9]*|(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?[#@!%&]?|[^\s]',text):
  t=m.group()
  if t.startswith("'"):continue
  if t.startswith('"'):result.append(t);continue
  if re.fullmatch(r'(?:\d+\.\d*|\.\d+|\d+)(?:[eE][-+]?\d+)?[#@!%&]?',t) and ('.' in t or 'e' in t.lower() or t.endswith('#')):
   try:t='DOUBLE:'+struct.pack('<d',float(t.rstrip('#'))).hex()
   except ValueError:pass
  result.append(t.lower())
 return result
def macros(path):
 p=VBA_Parser(str(path))
 try:return {name.rsplit('.',1)[0]:text for _,_,name,text in p.extract_macros()}
 finally:p.close()
def saved_cells(path):
 ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
 with zipfile.ZipFile(path) as z:
  shared=[''.join(e.itertext()) for e in ET.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',ns)] if 'xl/sharedStrings.xml' in z.namelist() else []
  rels={e.attrib['Id']:e.attrib['Target'] for e in ET.fromstring(z.read('xl/_rels/workbook.xml.rels'))};result={}
  for sheet in ET.fromstring(z.read('xl/workbook.xml')).findall('s:sheets/s:sheet',ns):
   target=rels[sheet.attrib['{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id']];target=target.lstrip('/') if target.startswith('/') else 'xl/'+target;cells={}
   for cell in ET.fromstring(z.read(target)).findall('.//s:sheetData/s:row/s:c',ns):
    v=cell.find('s:v',ns);typ=cell.get('t')
    if typ=='inlineStr':value=''.join(cell.find('s:is',ns).itertext())
    elif v is None or v.text is None:continue
    elif typ=='s':value=shared[int(v.text)]
    elif typ in ('str','e'):value=v.text
    else:value=float(v.text)
    cells[cell.attrib['r']]=value
   result[sheet.attrib['name']]=cells
  return result
def log(path):
 data=path.read_bytes();enc='utf-16' if data[:2] in [b'\xff\xfe',b'\xfe\xff'] else 'utf-8-sig'
 return list(csv.DictReader(data.decode(enc).splitlines(),delimiter='\t'))
def parts(text):return dict(x.split('=',1) for x in text.split(';') if '=' in x)
def compare_logs(a,b,max_solve_id=None):
 old=log(a);new=log(b);errors=[]
 if max_solve_id is not None:
  old=[r for r in old if int(r['solve_id'])<=max_solve_id];new=[r for r in new if int(r['solve_id'])<=max_solve_id]
 def solves(rows):
  starts=[(r['candidate'],r['mode'],r['strength_factor'],r['detail']) for r in rows if r['event']=='solve_start']
  keys=['status','last_iteration_index','objective','primal','dual','gap','message']
  ends=[tuple(parts(r['detail']).get(k) for k in keys) for r in rows if r['event']=='solve_end']
  return {'starts':starts,'ends':ends}
 sa,sb=solves(old),solves(new)
 if sa!=sb:errors.append({'solve_sequence':{'old':sa,'new':sb}})
 def audits(rows):return [r['detail'] for r in rows if r['event']=='audit']
 aa,ab=audits(old),audits(new)
 if aa!=ab:errors.append({'audit_differences':True,'old_count':len(aa),'new_count':len(ab)})
 return {'solves':len(sb['starts']),'completed_solves':len(sb['ends']),'audits':len(ab),'matches':not errors,'errors':errors,'old_sha256':sha(a),'new_sha256':sha(b),'maximum_compared_solve_id':max_solve_id}
def verify_book(path):
 baseline=json.loads((ROOT/'results/baseline.json').read_text(encoding='utf-8'));assert all(sha(REPO/p)==h for p,h in baseline['protected'].items())
 old=macros(SOURCE);new=macros(path);sources={p.stem:p.read_text(encoding='utf-8') for p in (ROOT/'src/vba').iterdir() if p.is_file()}
 baseline_sources={p.stem:p.read_text(encoding='utf-8') for p in (ROOT/'baseline_sources').iterdir() if p.is_file()}
 assert set(baseline_sources)==set(old) and all(canonical(v)==canonical(old[n]) for n,v in baseline_sources.items())
 assert baseline['source_sha256']==sha(SOURCE)
 assert set(new)==set(sources) and set(new)-set(old)=={'RPX_NativeSolver'}
 assert all(canonical(v)==canonical(new[n]) for n,v in sources.items())
 changed=[n for n in old if canonical(old[n])!=canonical(new[n])];assert set(changed)=={'RPX_HSDLu','RPX_LDL','RPX_Native','RPX_NativeBulk','RPX_NativeLU','RPX_Solve'}
 assert all(canonical(old[n])==canonical(new[n]) for n in old if n.startswith('Sheet') or n=='ThisWorkbook')
 imports=json.loads((ROOT/'results/import_modules.json').read_text(encoding='utf-8'))['modules']
 for name in imports:
  data=(ROOT/'import'/f'{name}.bas').read_bytes();assert data.decode('cp932').replace('\r\n','\n')==sources[name];assert b'\r\n' in data
 diff=[]
 for name in sorted(changed+['RPX_NativeSolver']):
  a=(ROOT/'baseline_sources'/f'{name}.bas').read_text(encoding='utf-8').splitlines(keepends=True) if name in old else []
  diff.extend(difflib.unified_diff(a,sources[name].splitlines(keepends=True),fromfile='G1R16_DLL_3/'+name+'.bas' if a else '/dev/null',tofile='G1R16_DLL4/'+name+'.bas'))
 (ROOT/'changes.diff').write_text(''.join(diff),encoding='utf-8',newline='\n')
 native_diff=[]
 for p in sorted((ROOT/'native').iterdir()):
  if not p.is_file() or p.suffix not in ['.c','.cpp','.h','.asm','.cmd']:continue
  old_path=DLL3/'native'/p.name
  a=old_path.read_text(encoding='utf-8').splitlines(keepends=True) if old_path.exists() else []
  b=p.read_text(encoding='utf-8').splitlines(keepends=True)
  native_diff.extend(difflib.unified_diff(a,b,fromfile='G1R16_DLL_3/native/'+p.name if a else '/dev/null',tofile='G1R16_DLL4/native/'+p.name))
 (ROOT/'native_changes.diff').write_text(''.join(native_diff),encoding='utf-8',newline='\n')
 return {'protected_original_files':len(baseline['protected']),'protected_hashes_match':True,'baseline_modules':len(old),'baseline_sources_match_actual_DLL3':True,'delivery_modules':len(new),'changed_existing_modules':sorted(changed),'added_modules':['RPX_NativeSolver'],'unchanged_existing_modules':len(old)-len(changed),'document_modules_unchanged':True,'sources_match_saved_workbook':True,'comparison':'case-insensitive identifiers; exact string literals and IEEE754-equivalent VBE numeric formatting; excludes comments/attributes','cp932_crlf_strict_roundtrip':True,'xlsm_sha256':sha(path),'dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll')}
if __name__=='__main__':
 import argparse
 p=argparse.ArgumentParser();p.add_argument('workbook',type=Path);args=p.parse_args();record=verify_book(args.workbook);write(ROOT/'results/delivery_verification.json',record);print(json.dumps(record,ensure_ascii=False))
