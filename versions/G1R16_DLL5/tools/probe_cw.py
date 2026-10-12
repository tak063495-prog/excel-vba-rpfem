import ctypes as C,shutil,win32com.client,json,gc
from prepare import ROOT,SOURCE,write
from native_helpers import compile_book
folder=ROOT/'results/nt_oracle'
folder.mkdir(parents=True,exist_ok=True)
book_path=folder/'NT_Oracle_G1R16.xlsm'
if not book_path.exists():
    book_path=folder/'CW_Oracle_G1R16.xlsm'
    shutil.copyfile(SOURCE,book_path)
shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
app=win32com.client.DispatchEx('Excel.Application');book=None
try:
    app.Visible=False;app.DisplayAlerts=False
    book=app.Workbooks.Open(str(book_path),UpdateLinks=0)
    code='''Option Explicit
Private Declare PtrSafe Function LL Lib "kernel32" Alias "LoadLibraryExW" (ByVal path As LongPtr, ByVal f As LongPtr, ByVal flags As Long) As LongPtr
Private Declare PtrSafe Function CW Lib "rpfem_native_x64.dll" Alias "rpx_test_controlword" (ByVal mode As Long, ByVal word As Long) As Long
Public Function GetCW(ByVal path As String) As Long
    Dim h As LongPtr
    h = LL(StrPtr(path), 0, &H1100)
    If h = 0 Then err.Raise 513, , "DLL load failed"
    GetCW = CW(0,0)
End Function
'''
    comp=book.VBProject.VBComponents.Add(1);comp.Name='CWOracle';comp.CodeModule.AddFromString(code)
    compile_book(app,book)
    cw=int(app.Run("'"+book.Name+"'!GetCW",str(folder/'rpfem_native_x64.dll')))
    dll=C.CDLL(str(ROOT/'delivery/rpfem_native_x64.dll'));fn=dll.rpx_test_controlword;fn.argtypes=[C.c_int,C.c_int];fn.restype=C.c_int
    write(folder/'controlword.json',{'excel_x87':cw,'python_x87':fn(0,0)})
    print('CONTROL WORD EXCEL',hex(cw),'PYTHON',hex(fn(0,0)),flush=True)
finally:
    if book is not None:book.Close(False)
    app.Quit();book=None;app=None;gc.collect()
