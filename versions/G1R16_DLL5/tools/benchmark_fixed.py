"""Actual native Excel, three backends, same fixed Fs and deterministic q8 mesh.
VBA timing here is incidental to accuracy checks; no full VBA Fs search is repeated.
"""
import argparse,gc,json,shutil,struct,time
import numpy as np
import win32com.client
from prepare import ROOT,REPO,DLL3,DLL4,write,sha
from native_helpers import compile_book
from test_hsd_regression import inject,setvalue
from test_fem_precision import mesh_from_files,physical_audit

def configure(book,rec,target,q,diagnostic=False):
 app=book.Application;app.EnableEvents=False
 ws=book.Worksheets('要素定義');ws.Range('A7:T1006').ClearContents()
 for first,last,key in [('A','C','points'),('E','H','regions'),('J','T','faces')]:
  ws.Range(first+'7:'+last+str(6+len(rec[key]))).Value2=tuple(tuple(x) for x in rec[key])
 ws=book.Worksheets('材料データ');ws.Range('A5:Q1004').ClearContents();ws.Range('A5:E'+str(4+len(rec['materials']))).Value2=tuple(tuple(x) for x in rec['materials'])
 for key,value in {'GRAVITY':'固定','TARGET':target,'QUADRATURE':q,'ADAPT':0,'HSD_CAPTURE':0,'TEST_INJECTION':0,'HSD_SOLVE_SECONDS':7200,'HSD_LINEAR_BACKEND':'AUTO_PIVOT','ANALYSIS_POLICY':rec['settings']['ANALYSIS_POLICY'],'SEISMIC_KH':0}.items():setvalue(book,key,value)
 book.Worksheets('浸透入力').Range('D4').Value2=0;app.EnableEvents=True
 if diagnostic:setvalue(book,'DIAGNOSTIC_MODE',1)

def run():
 p=argparse.ArgumentParser();p.add_argument('--tag',required=True);p.add_argument('--target',type=int,default=256);p.add_argument('--q',type=int,default=8);p.add_argument('--diagnostic',action='store_true')
 p.add_argument('--cases',nargs='+',choices=['C5_45','P06A','P06N'],default=['C5_45','P06A','P06N']);p.add_argument('--backends',nargs='+',choices=['vba','dll3','dll4','dll5'],default=['vba','dll4','dll5']);p.add_argument('--sides',nargs='+',choices=['upper','lower'],default=['upper','lower']);args=p.parse_args();assert args.tag.isidentifier() and len(set(args.backends))==len(args.backends)
 folder=ROOT/'results'/args.tag;folder.mkdir(parents=True,exist_ok=False)
 result={'scope':'Actual Excel paired fixed-Fs solves, three backends; not final safety-factor search','target':args.target,'q':args.q,'timing_method':'serial fresh Excel process per backend; one solve per case and side; physical audit excluded from solve timing','cases':{},'backends':{}}
 for backend in args.backends:
  kit=DLL3 if backend=='dll3' else DLL4 if backend in ['dll4','vba'] else ROOT
  filename='RPFEM_20261012_G1R16_DLL_3_Levee_Checked.xlsm' if backend=='dll3' else 'RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm' if backend in ['dll4','vba'] else 'RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm'
  source=kit/'delivery'/filename
  sub=folder/backend;sub.mkdir();path=sub/'FixedBenchmark.xlsm';shutil.copyfile(source,path);shutil.copyfile(kit/'delivery/rpfem_native_x64.dll',sub/'rpfem_native_x64.dll')
  result['backends'][backend]={'source_sha256':sha(source),'dll_sha256':sha(sub/'rpfem_native_x64.dll')}
  app=win32com.client.DispatchEx('Excel.Application');book=None
  try:
   app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
   for name in ['HSDRegression','MeshProbe']:inject(book,name)
   compile_book(app,book);call=lambda name,*a:app.Run("'"+book.Name+"'!"+name,*a)
   call('RPX_NativeSetMode','vba' if backend=='vba' else 'auto')
   for case in args.cases:
    rec=json.loads((REPO/'validation/G1R12_native25_20261009/inputs'/('P01A.json' if case=='C5_45' else case+'.json')).read_text(encoding='utf-8'))
    if case=='C5_45':
     rec['materials']=[[1,5.,45.,20.,45.]];rec['source_case']['materials'][0].update(c=5.,phi=45.,gamma=20.,psi=45.);rec['source_case']['policy']='equal_reduced_phi'
    configure(book,rec,args.target,args.q,args.diagnostic)
    for side in args.sides:
     key=case+'_'+side;fs=3. if case=='C5_45' else 1.;prefix=sub/key
     print('START',backend,key,flush=True);t=time.perf_counter();message=str(call('FixedFem',side,fs,str(prefix)));seconds=time.perf_counter()-t
     row={'Fs_input':fs,'message':message,'seconds':seconds};result['cases'].setdefault(key,{})[backend]=row;write(folder/'summary.json',result);assert message.startswith('PASS'),message
     raw=prefix.with_name(prefix.name+'_field.bin').read_bytes();n=struct.unpack_from('<i',raw,8)[0];x=np.frombuffer(raw,'<f8',n,12).copy();m=mesh_from_files(prefix,rec)
     row['elements']=len(m.tri);row['independent_audit']=physical_audit(m,x[:12*len(m.tri)] if side=='upper' else x,side,fs,args.q);assert row['independent_audit']['pass']
     row['field_sha256']=sha(prefix.with_name(prefix.name+'_field.bin'));row['mesh_sha256']=sha(prefix.with_name(prefix.name+'_mesh.bin'))
     write(folder/'summary.json',result);print('DONE',backend,key,seconds,message,flush=True)
  finally:
   if book is not None:book.Close(False)
   app.Quit();book=None;app=None;gc.collect()
 for row in result['cases'].values():
  row['bitwise_exact_backends']=len({row[b]['field_sha256'] for b in args.backends})==1 and len({row[b]['mesh_sha256'] for b in args.backends})==1
  row['compared_backends']=args.backends
  if 'dll4' in row and 'dll3' in row:row['reduction_percent_vs_DLL3']=100*(1-row['dll4']['seconds']/row['dll3']['seconds'])
  if 'dll5' in row and 'dll4' in row:row['reduction_percent_vs_DLL4']=100*(1-row['dll5']['seconds']/row['dll4']['seconds'])
  assert row['bitwise_exact_backends'],row
 result['passed']=True;write(folder/'summary.json',result);print('BACKEND q8 FEM PASS',len(result['cases']),args.backends,flush=True)
if __name__=='__main__':run()
