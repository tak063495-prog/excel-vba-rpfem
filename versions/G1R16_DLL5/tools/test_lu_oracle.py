import gc,json,shutil,struct,time
import numpy as np
import win32com.client
from prepare import ROOT,write,sha
from native_helpers import compile_book
def run():
    folder=ROOT/'results/lu_oracle';folder.mkdir(parents=True,exist_ok=True)
    rng=np.random.default_rng(16003);cases=[]
    for n in [1,2,3,7,19,47,97]:
        for sym in [0,1]:
            a=rng.normal(size=(n,n));a[np.abs(a)<.6]=0
            if sym:a=np.triu(a);a=a+a.T-np.diag(a.diagonal())
            a[np.diag_indices(n)]+=n+1
            for eq in [0,1]:
                for stream in [0,1]:
                    for threshold in [0.,.2,1.]:
                        scale=np.geomspace(1e-6,1e6,n);matrix=a*scale[:,None]*scale[None,:]
                        cases.append((matrix,sym,eq,stream,threshold,rng.permutation(n),rng.normal(size=n)))
    # Forced pivot ties, physical row swaps, exact entry deletion, singular return.
    for matrix in [np.array([[0.,2.,1.],[2.,0.,4.],[2.,3.,2.]]),np.ones((4,4)),np.array([[1.,1.],[1.,2.]])]:
        for eq in [0,1]:
            for stream in [0,1]:cases.append((matrix,0,eq,stream,0.,np.arange(len(matrix)),np.arange(len(matrix))+1.))
    cases=[(*case,budget) for case in cases for budget in [1e9,1e8]]
    parts=[struct.pack('<i',len(cases))]
    for a,sym,eq,stream,th,perm,rhs,budget in cases:
        n=len(a);ap=[0];ai=[];av=[]
        for c in range(n):
            for r in range(c+1 if sym else n):
                if a[r,c]!=0:ai.append(r);av.append(a[r,c])
            ap.append(len(ai))
        parts.extend([struct.pack('<5i2d',n,len(ai),sym,eq,stream,th,budget),np.asarray(perm,dtype='<i4').tobytes(),np.asarray(ap,dtype='<i4').tobytes(),np.asarray(ai,dtype='<i4').tobytes(),np.asarray(av,dtype='<f8').tobytes(),np.asarray(rhs,dtype='<f8').tobytes()])
    (folder/'inputs.bin').write_bytes(b''.join(parts))
    path=folder/'LU_Oracle.xlsm';shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R16_DLL5_Levee_Checked.xlsm',path)
    shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
    app=win32com.client.DispatchEx('Excel.Application');book=None;metadata={}
    try:
        app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
        cm=book.VBProject.VBComponents('RPX_HSDLu').CodeModule
        cm.AddFromString((ROOT/'tests/LUInstrumentation.txt').read_text(encoding='utf-8'))
        comp=book.VBProject.VBComponents.Add(1);comp.Name='LUOracle'
        comp.CodeModule.AddFromString('\r\n'.join(s for s in (ROOT/'tests/LUOracle.bas').read_text(encoding='utf-8').splitlines() if not s.startswith('Attribute VB_')))
        compile_book(app,book)
        for mode in ['vba','auto']:
            t=time.perf_counter();msg=app.Run("'"+book.Name+"'!DumpLUOracle",str(folder/'inputs.bin'),str(folder/(mode+'.bin')),mode)
            metadata[mode]={'message':msg,'seconds':time.perf_counter()-t};print(mode,msg,flush=True)
        book.Save()
    finally:
        if book is not None:book.Close(False)
        app.Quit();book=None;app=None;gc.collect()
    old=(folder/'vba.bin').read_bytes();new=(folder/'auto.bin').read_bytes();pos=[4,4];failures=[];passed=singular=0;residuals=[]
    for index,(a,sym,eq,stream,th,perm,rhs,budget) in enumerate(cases):
        outputs=[]
        for j,data in enumerate([old,new]):
            start=pos[j];eno,accepted=struct.unpack_from('<ii',data,start);pos[j]+=8;x=None
            if eno==0 and accepted:
                x=np.frombuffer(data,dtype='<f8',count=len(a),offset=pos[j]).copy();pos[j]+=8*(len(a)+3)
                lc,uc=struct.unpack_from('<ii',data,pos[j]);pos[j]+=8+len(a)*32+(len(a)+1)*8+(lc+uc)*12
            outputs.append((eno,accepted,data[start:pos[j]],x))
        b,c=outputs
        if b[2]!=c[2]:failures.append({'index':index,'n':len(a),'vba_error':b[0],'native_error':c[0],'vba_accepted':b[1],'native_accepted':c[1],'max_answer_difference':float(np.max(abs(b[3]-c[3]))) if b[3] is not None and c[3] is not None else None})
        elif b[0]:failures.append({'index':index,'both_error':b[0]})
        elif b[1]:
            passed+=1;r=np.linalg.norm(a@b[3]-rhs,np.inf)/(np.linalg.norm(a,np.inf)*np.linalg.norm(b[3],np.inf)+np.linalg.norm(rhs,np.inf));residuals.append(float(r))
        else:singular+=1
    assert pos==[len(old),len(new)]
    result={'cases':len(cases),'exact_factor_and_solution':passed,'singular_identical':singular,'failures':failures,'max_independent_backward_error':max(residuals),'timing':metadata,'dll_sha256':sha(ROOT/'delivery/rpfem_native_x64.dll')}
    write(folder/'summary.json',result);print(json.dumps(result,ensure_ascii=False),flush=True)
    assert not failures and max(residuals)<1e-12
if __name__=='__main__':run()
