from pathlib import Path
import zipfile,xml.etree.ElementTree as E,json,hashlib
R=Path(__file__).resolve().parent
P=Path('C:/Users/link_/Desktop/RPFEM_G1R9_Speedup_20261006/RPFEM_20261006_G1R9_P06N_DD_Fast_Verified.xlsm')
z=zipfile.ZipFile(P);ns={'s':'http://schemas.openxmlformats.org/spreadsheetml/2006/main'}
ss=[''.join(t.itertext()) for t in E.fromstring(z.read('xl/sharedStrings.xml')).findall('s:si',ns)]
def cells(n):
 out={}
 for c in E.fromstring(z.read(n)).findall('.//s:sheetData/s:row/s:c',ns):
  v=c.find('s:v',ns)
  if v is not None:out[c.get('r')]=ss[int(v.text)] if c.get('t')=='s' else v.text
 return out
if __name__=='__main__':
 for n in ['xl/worksheets/sheet7.xml','xl/worksheets/sheet8.xml','xl/worksheets/sheet5.xml']:
  d=cells(n);print(n,[(k,v) for k,v in d.items() if int(''.join(c for c in k if c.isdigit()))<=2])
 (R/'book_identity.json').write_text(json.dumps({'path':str(P),'sha256':hashlib.sha256(P.read_bytes()).hexdigest()},indent=2))
