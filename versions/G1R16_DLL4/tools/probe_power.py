import ctypes as C,struct,json,win32com.client,gc,math
from prepare import ROOT,write
from native_helpers import compile_book
folder=ROOT/'results/nt_oracle'
if not (folder/'powers.bin').exists():
    app=win32com.client.DispatchEx('Excel.Application');book=None
    try:
        app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(folder/'NT_Oracle_G1R16.xlsm'),UpdateLinks=0)
        comp=book.VBProject.VBComponents.Add(1);comp.Name='PowerOracle'
        comp.CodeModule.AddFromString('\r\n'.join(s for s in (ROOT/'tests/PowerOracle.bas').read_text(encoding='utf-8').splitlines() if not s.startswith('Attribute VB_')))
        compile_book(app,book)
        print(app.Run("'"+book.Name+"'!DumpPowerOracle",str(folder/'inputs.bin'),str(folder/'powers.bin')),flush=True)
    finally:
        if book is not None:book.Close(False)
        app.Quit();book=None;app=None;gc.collect()
raw=(folder/'inputs.bin').read_bytes();n=struct.unpack_from('<i',raw)[0];inp=4
expected=(folder/'powers.bin').read_bytes();pos=4
dll=C.CDLL(str(ROOT/'delivery/rpfem_native_x64.dll'));dll.rpx_test_controlword(1,json.loads((folder/'controlword.json').read_text(encoding='utf-8'))['excel_x87'])
fn=dll.rpx_nt_scalar_test;fn.argtypes=[C.c_int32,C.POINTER(C.c_double),C.c_double,C.POINTER(C.c_double)]
diffs=[]
for k in range(n):
    row=struct.unpack_from('<i6d',raw,inp);inp+=52
    for side in range(2):
        a=(C.c_double*3)(*row[1+side*3:4+side*3]);rad=struct.unpack_from('<d',expected,pos)[0];pos+=8
        for pw in [.5,-.5,-1.]:
            eno=struct.unpack_from('<i',expected,pos)[0];ex=struct.unpack_from('<3d',expected,pos+4);pos+=28
            out=(C.c_double*3)();code=fn(0,a,pw,out)
            if (eno==0)!=(code==0) or (eno==0 and any(struct.pack('<d',x)!=struct.pack('<d',y) for x,y in zip(ex,out))):
                diffs.append({'case':k,'side':side,'power':pw,'a':list(a),'radius':rad,'vba_error':eno,'native_error':code,'vba':ex,'native':list(out)})
        ex=struct.unpack_from('<9d',expected,pos);pos+=72;out=(C.c_double*9)();fn(1,a,0,out)
        if any(struct.pack('<d',x)!=struct.pack('<d',y) for x,y in zip(ex,out)):diffs.append({'case':k,'side':side,'quad':True,'vba':ex,'native':list(out)})
write(folder/'power_comparison.json',{'differences':diffs,'rows':n*2*4})
print('POWER/QUAD DIFFERENCES',len(diffs),'FIRST',json.dumps(diffs[:5]),flush=True)
