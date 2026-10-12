import gc,copy,json,math,shutil,struct,sys,time,collections,argparse
import numpy as np
import win32com.client
from prepare import ROOT,REPO,write,sha
from native_helpers import compile_book
from test_hsd_regression import inject,setvalue
sys.path.insert(0,str(REPO/'versions/G1R12/validation_runtime/python'))
from common import Mesh,physical_audit
def mesh_from_files(prefix,rec):
    data=prefix.with_name(prefix.name+'_mesh.bin').read_bytes();nn,ne=struct.unpack_from('<ii',data);pos=8
    xy=np.frombuffer(data,'<f8',nn*2,pos).copy().reshape(nn,2);pos+=nn*16
    tr=np.frombuffer(data,'<i4',ne*4,pos).copy().reshape(ne,4);assert pos+ne*16==len(data)
    mats=np.asarray([[m['c'],m['phi'],m['psi'],m['gamma']] for m in rec['source_case']['materials']])
    edges=collections.Counter(tuple(sorted((int(a),int(b)))) for t in tr for a,b in zip(t[:3],np.roll(t[:3],-1)))
    bd={};tol=1e-8
    def on(p,a,b):
        a=np.asarray(a);b=np.asarray(b);d=b-a;t=(p-a)@d/(d@d)
        return -tol<=t<=1+tol and abs(d[0]*(p[1]-a[1])-d[1]*(p[0]-a[0]))<tol*np.linalg.norm(d)
    kinds={'自由':'load','固定XY':'fixed','固定X':'roller_x','固定Y':'roller_y'}
    for edge,num in edges.items():
        if num!=1:continue
        matches=[f for f in rec['boundary_explicit'] if all(on(xy[i],f['a'],f['b']) for i in edge)];assert len(matches)==1
        f=matches[0];loads=np.zeros((2,2));loads[:,1]=np.asarray(f['load0'])+np.asarray(f['load1'])
        bd[edge]={'kind':kinds[f['kind']],'rigid':-1,'loads':loads}
    body=np.zeros((ne,2,2));body[:,1,1]=-mats[tr[:,3],3]
    return Mesh(xy,tr[:,:3],mats,tr[:,3],np.full(ne,-1),body,bd,max(1.,float(np.max(mats[:,0]))),'dll3_fixed',load_phase='reference_total',provenance={'psi_policy':rec['source_case']['policy']}).validate()
def run():
    parser=argparse.ArgumentParser();parser.add_argument('--tag',default='fixed_fem_verified');args=parser.parse_args();assert args.tag.isidentifier()
    folder=ROOT/'results'/args.tag;folder.mkdir(parents=True,exist_ok=False)
    path=folder/'Fixed_FEM.xlsm';shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm',path)
    shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
    app=win32com.client.DispatchEx('Excel.Application');book=None;result={'scope':'same fixed Fs, deterministic TARGET64 mesh, q2, dry; not full adaptive P06 Fs search','cases':[],'dll_sha256':sha(folder/'rpfem_native_x64.dll')}
    try:
        app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
        for name in ['HSDRegression','MeshProbe']:inject(book,name)
        compile_book(app,book);call=lambda name,*args:app.Run("'"+book.Name+"'!"+name,*args)
        for case in ['C5_45','P06A','P06N']:
            rec=json.loads((REPO/'validation/G1R12_native25_20261009/inputs'/('P01A.json' if case=='C5_45' else case+'.json')).read_text(encoding='utf-8'))
            if case=='C5_45':
                rec['materials']=[[1,5.,45.,20.,45.]]
                rec['source_case']['materials'][0].update(c=5.,phi=45.,gamma=20.,psi=45.)
                rec['source_case']['policy']='equal_reduced_phi'
            app.EnableEvents=False
            ws=book.Worksheets('要素定義');ws.Range('A7:T1006').ClearContents()
            for first,last,key in [('A','C','points'),('E','H','regions'),('J','T','faces')]:
                values=tuple(tuple(x) for x in rec[key]);ws.Range(first+'7:'+last+str(6+len(values))).Value2=values
            ws=book.Worksheets('材料データ');ws.Range('A5:Q1004').ClearContents();ws.Range('A5:E'+str(4+len(rec['materials']))).Value2=tuple(tuple(x) for x in rec['materials'])
            for key,val in {'TARGET':64,'QUADRATURE':2,'ADAPT':0,'HSD_CAPTURE':0,'TEST_INJECTION':0,'HSD_SOLVE_SECONDS':120,'HSD_LINEAR_BACKEND':'AUTO_PIVOT','ANALYSIS_POLICY':rec['settings']['ANALYSIS_POLICY'],'SEISMIC_KH':0}.items():setvalue(book,key,val)
            book.Worksheets('浸透入力').Range('D4').Value2=0
            app.EnableEvents=True
            for side in ['upper','lower']:
                record={'case':case,'side':side,'Fs':3.0 if case=='C5_45' else 1.0,'passes':[]}
                for mode in ['vba','auto']:
                    prefix=folder/f'{case}_{side}_{mode}';call('RPX_NativeSetMode',mode)
                    t=time.perf_counter();msg=str(call('FixedFem',side,record['Fs'],str(prefix)))
                    record[mode]={'message':msg,'seconds':time.perf_counter()-t};print(case,side,mode,msg,flush=True)
                    assert msg.startswith('PASS'),msg
                    raw=prefix.with_name(prefix.name+'_field.bin').read_bytes();assert raw[:8]==b'P06X0001';n=struct.unpack_from('<i',raw,8)[0];x=np.frombuffer(raw,'<f8',n,12).copy()
                    m=mesh_from_files(prefix,rec);audit=physical_audit(m,x[:12*len(m.tri)] if side=='upper' else x,side,record['Fs'],2);record[mode]['independent_audit']=audit;assert audit['pass'],audit
                record['same_mesh']= (folder/f'{case}_{side}_vba_mesh.bin').read_bytes()==(folder/f'{case}_{side}_auto_mesh.bin').read_bytes()
                record['bitwise_same_field']= (folder/f'{case}_{side}_vba_field.bin').read_bytes()==(folder/f'{case}_{side}_auto_field.bin').read_bytes()
                assert record['same_mesh'] and record['bitwise_same_field']
                result['cases'].append(record);write(folder/'summary.json',result)
    finally:
        write(folder/'summary.json',result)
        if book is not None:book.Close(False)
        app.Quit();book=None;app=None;gc.collect()
    print('FEM precision PASS',len(result['cases']),flush=True)
if __name__=='__main__':run()
