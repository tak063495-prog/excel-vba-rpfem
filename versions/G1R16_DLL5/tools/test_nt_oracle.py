import ctypes as C,struct,json,math,shutil,random,time,gc,sys,subprocess
import win32com.client,win32process,pythoncom
from prepare import ROOT,SOURCE,sha,write
from native_helpers import compile_book
class DD(C.Structure):_fields_=[('hi',C.c_double),('lo',C.c_double)]
def run():
    extended='--extended' in sys.argv
    folder=ROOT/('results/nt_oracle_extended' if extended else 'results/nt_oracle');folder.mkdir(parents=True,exist_ok=True)
    rng=random.Random(16003);cases=[]
    for precise in [0,1]:
        for scale in [1e-8,1e-3,1.,1e3,1e8]:
            for margin in [.8,.1,1e-4,1e-7,1e-10,1e-13]:
                for _ in range(3):
                    a=rng.uniform(-math.pi,math.pi);b=rng.uniform(-math.pi,math.pi)
                    s=[scale,scale*(1-margin)*math.cos(a),scale*(1-margin)*math.sin(a)]
                    z=[scale*1.7,scale*1.7*(1-margin)*math.cos(b),scale*1.7*(1-margin)*math.sin(b)]
                    cases.append((precise,s,z))
        cases += [(precise,[1.,0.,0.],[2.,0.,0.]),(precise,[1.,1.,0.],[1.,0.,0.]),(precise,[0.,0.,0.],[1.,0.,0.]),(precise,[-1.,0.,0.],[1.,0.,0.])]
        if extended:
            for scale in [1e-150,1e-140,1e-100,1e100,1e140,1e150]:
                for margin in [.8,.1,1e-6]:
                    cases.append((precise,[scale,scale*(1-margin),0.],[scale*1.7,0.,scale*1.7*(1-margin)]))
    data=struct.pack('<i',len(cases))+b''.join(struct.pack('<i6d',p,*s,*z) for p,s,z in cases)
    (folder/'inputs.bin').write_bytes(data)
    book_path=folder/'NT_Oracle_G1R16.xlsm'
    if not (folder/'vba.bin').exists():
        assert not book_path.exists();shutil.copyfile(SOURCE,book_path)
        app=win32com.client.DispatchEx('Excel.Application');book=None
        try:
            app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(book_path),UpdateLinks=0)
            comp=book.VBProject.VBComponents.Add(1);comp.Name='NTOracle'
            text=(ROOT/'tests/NTOracle.bas').read_text(encoding='utf-8')
            comp.CodeModule.AddFromString('\r\n'.join(s for s in text.splitlines() if not s.startswith('Attribute VB_')))
            compile_book(app,book)
            t=time.perf_counter();result=app.Run("'"+book.Name+"'!DumpNTOracle",str(folder/'inputs.bin'),str(folder/'vba.bin'))
            write(folder/'oracle_metadata.json',{'message':result,'seconds':time.perf_counter()-t,'excel_pid':win32process.GetWindowThreadProcessId(app.Hwnd)[1],'security_unmodified':app.AutomationSecurity,'baseline_sha256':sha(SOURCE)})
            print(result,flush=True);book.Save()
        finally:
            if book is not None:book.Close(False)
            app.Quit();book=None;app=None;gc.collect()
    dll=C.CDLL(str(ROOT/'delivery/rpfem_native_x64.dll'))
    cw_path=ROOT/'results/nt_oracle/controlword.json'
    # Match the observed Excel thread only in this Python comparison process.
    # The production DLL must never change the caller's floating-point mode.
    if not cw_path.exists():
        subprocess.run([sys.executable,str(ROOT/'tools/probe_cw.py')],check=True)
    if cw_path.exists():
        control=dll.rpx_test_controlword;control.argtypes=[C.c_int,C.c_int];control.restype=C.c_int
        control(1,json.loads(cw_path.read_text(encoding='utf-8'))['excel_x87'])
    fp=C.POINTER(C.c_double);ddp=C.POINTER(DD)
    fn=dll.rpx_nt_test;fn.argtypes=[C.c_int32]+[fp]*7+[ddp];fn.restype=C.c_int
    raw=(folder/'vba.bin').read_bytes();assert struct.unpack_from('<i',raw)[0]==len(cases)
    pos=4;errors=[];summary={'rows':len(cases),'normal_success':0,'precise_success':0,'bitwise_exact_rows':0,'maximum_relative_error':0.,'differences':[]}
    for index,(p,s,z) in enumerate(cases):
        eno=struct.unpack_from('<i',raw,pos)[0];pos+=4
        expected=struct.unpack_from('<51d',raw,pos);pos+=51*8
        a=(C.c_double*3)(*s);b=(C.c_double*3)(*z);h=(C.c_double*9)();wm=(C.c_double*9)();wi=(C.c_double*9)();si=(C.c_double*3)();v=(C.c_double*3)();hd=(DD*9)()
        code=fn(p,a,b,h,wm,wi,si,v,hd)
        out=list(h)+list(wm)+list(wi)+list(si)+list(v)+[x for d in hd for x in (d.hi,d.lo)]
        mapped=0 if code==0 else 6 if code==206 else 11 if code==205 else 5
        if eno!=mapped:errors.append({'index':index,'precise':p,'vba_error':eno,'native_code':code,'s':s,'z':z})
        if eno or code:continue
        summary['precise_success' if p else 'normal_success']+=1
        diffs=[j for j,(x,y) in enumerate(zip(expected,out)) if struct.pack('<d',x)!=struct.pack('<d',y)]
        if not diffs:summary['bitwise_exact_rows']+=1
        else:
            rel=max(abs(expected[j]-out[j])/max(abs(expected[j]),abs(out[j]),1e-300) for j in diffs)
            summary['maximum_relative_error']=max(summary['maximum_relative_error'],rel)
            summary['differences'].append({'index':index,'precise':p,'relative_error':rel,'components':diffs,'first':{'component':diffs[0],'vba':expected[diffs[0]],'native':out[diffs[0]]}})
    summary['classification_mismatches']=errors
    summary['dll_sha256']=sha(ROOT/'delivery/rpfem_native_x64.dll')
    write(folder/'comparison.json',summary)
    print(json.dumps({k:v for k,v in summary.items() if k not in ['differences','classification_mismatches']},ensure_ascii=False),flush=True)
    print('DIFFERING ROWS',len(summary['differences']),'CLASSIFICATION MISMATCHES',len(errors),flush=True)
    if errors:print(json.dumps(errors[:3]),flush=True)
    if summary['differences']:print(json.dumps(summary['differences'][:4]),flush=True)
    assert not errors and not summary['differences']
if __name__=='__main__':run()
