import csv,difflib,json,re,struct,zipfile,xml.etree.ElementTree as ET
from pathlib import Path
from oletools.olevba import VBA_Parser
from support import sha
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
