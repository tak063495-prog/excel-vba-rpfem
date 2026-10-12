"""25 dry cases: mesh generation ONLY, no Fs solves or full analysis."""
import gc,json,shutil,struct,sys,numpy as np
import win32com.client
from support import ROOT,BASE,REPO,SOURCE,OUTPUT,write,sha,configure
from native_helpers import compile_book
from fixed_native import inject
def arrays(prefix):
 raw=prefix.with_name(prefix.name+'_mesh.bin').read_bytes();nn,ne=struct.unpack_from('<ii',raw)
 xy=np.frombuffer(raw,'<f8',2*nn,8).copy().reshape(nn,2);t=np.frombuffer(raw,'<i4',4*ne,8+16*nn).copy().reshape(ne,4)
 raw=prefix.with_name(prefix.name+'_physical.bin').read_bytes();assert struct.unpack_from('<i',raw)[0]==ne
 d=np.frombuffer(raw,np.dtype([('rigid','<i4'),('area','<f8'),('body','<f8',(4,))]),ne,8)
 masses=np.asarray([d['area'][t[:,3]==k].sum()for k in range(t[:,3].max()+1)])
 loads=(d['area'][:,None]*d['body']).sum(axis=0)
 return nn,ne,xy,t,masses,loads
def main():
 folder=ROOT/'results/geometry25'
 if '--verify-only'in sys.argv:
  meta=json.loads((folder/'summary.json').read_text(encoding='utf8'));assert len(meta['exports'])==75;meta['checks']=[]
 else:
  folder.mkdir(exist_ok=False);meta={'scope':'25 cases geometry only, TARGET200 q8 ADAPT0 dry kh0; no full runs','exports':[],'checks':[]}
 for label,source in ([]if '--verify-only'in sys.argv else [('old',SOURCE),('new',ROOT/'delivery'/OUTPUT)]):
  path=folder/(label+'.xlsm');shutil.copyfile(source,path)
  # New G1R17 is loaded only for the new book; old book uses forced VBA.
  dll=folder/'rpfem_native_x64.dll'
  if not dll.exists():shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',dll)
  app=win32com.client.DispatchEx('Excel.Application');book=None
  try:
   app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
   for name in ['MeshProbe','GeometryRegression']:inject(book,name)
   compile_book(app,book);call=lambda n,*a:app.Run("'"+book.Name+"'!"+n,*a)
   for inp in sorted((ROOT/'data/inputs').glob('*.json')):
    rec=json.loads(inp.read_text(encoding='utf8'));configure(book,rec)
    for mode in (['vba']if label=='old'else['vba','auto']):
     call('RPX_NativeSetMode',mode);prefix=folder/(inp.stem+'_'+label+'_'+mode)
     msg=str(call('GeometryExport',str(prefix)));assert msg.startswith('PASS'),msg
     assert prefix.with_name(prefix.name+'_mesh.bin').read_bytes()==prefix.with_name(prefix.name+'_repeat_mesh.bin').read_bytes()
     meta['exports'].append({'case':inp.stem,'version':label,'mode':mode,'result':msg,'idempotent':True});write(folder/'summary.json',meta)
     print(inp.stem,label,mode,msg.split(';DLL')[0],flush=True)
     if inp.stem in ['P06N','P06A'] and mode=='vba':
      msg=str(call('ExportCompiledPoint','lower',.5,str(folder/(inp.stem+'_'+label+'_model'))));assert msg.startswith('PASS'),msg
  finally:
   if book is not None:book.Close(False)
   app.Quit();book=None;app=None;gc.collect()
 for inp in sorted((ROOT/'data/inputs').glob('*.json')):
  key=inp.stem;old=folder/(key+'_old_vba');new=folder/(key+'_new_vba');native=folder/(key+'_new_auto')
  for suffix in ['_mesh.bin','_physical.bin','_edges.txt']:
   assert new.with_name(new.name+suffix).read_bytes()==native.with_name(native.name+suffix).read_bytes(),(key,suffix)
  a=arrays(old);b=arrays(new);assert np.array_equal(a[2],b[2][:a[0]])
  assert np.max(abs(a[4]-b[4]))<1e-9 and np.max(abs(a[5]-b[5]))<1e-8,(key,a[4:],b[4:])
  def semantics(path):
   result=[]
   for line in path.read_text(encoding='utf16').splitlines():
    fields=line.split('\t')
    if fields[2]=='interface':
     # Adjacent element encounter order can reverse after a split. Preserve
     # the two (material,rigid) sides and node pair, not side labels 0/1.
     sides=sorted([(fields[3],fields[5]),(fields[4],fields[6])]);fields=fields[:3]+[str(sides)]
    result.append(tuple(fields))
   return sorted(result)
  assert semantics(old.with_name(old.name+'_edges.txt'))==semantics(new.with_name(new.name+'_edges.txt')),key
  meta['checks'].append({'case':key,'old_nodes':a[0],'new_nodes':b[0],'old_elements':a[1],'new_elements':b[1],'vba_dll_mesh_loads_bitwise':True,'original_nodes_exact':True,'regions_and_raw_loads_preserved':True,'boundary_interfaces_exact':True,'unchanged_mesh':old.with_name(old.name+'_mesh.bin').read_bytes()==new.with_name(new.name+'_mesh.bin').read_bytes()})
 meta['PASS']=True;write(folder/'summary.json',meta);print('GEOMETRY 25 CASES PASS',flush=True)
if __name__=='__main__':main()
