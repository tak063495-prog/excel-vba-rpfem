"""Isolated Excel instances only. Test probes never enter the production workbook."""
from pathlib import Path
import sys, json, hashlib, shutil, time, os
import win32com.client
from native_helpers import compile_book
ROOT=Path(__file__).resolve().parents[1]
BASE=json.loads((ROOT/'source_identity.json').read_text(encoding='utf-8'))
candidate=Path(os.environ.get('RPFEM_BASELINE_XLSM',BASE['path']))
if not candidate.exists():
    candidate=ROOT.parent/'G1R11'/Path(BASE['path']).name
BASE['path']=str(candidate.resolve())
CHANGED=['RPX_Control','RPX_Output','RPX_Robust']
def replace(book,name,text):
    cm=book.VBProject.VBComponents(name).CodeModule
    cm.DeleteLines(1,cm.CountOfLines)
    cm.AddFromString('\r\n'.join(line for line in text.splitlines() if not line.startswith('Attribute VB_')))
def append(book,name,text):
    cm=book.VBProject.VBComponents(name).CodeModule
    cm.AddFromString(text)
def setting(book,key,value):
    ws=book.Worksheets('設定')
    for row in range(4,ws.UsedRange.Rows.Count+1):
        if ws.Cells(row,4).Value == key: ws.Cells(row,2).Value=value; return
    raise AssertionError('setting not found: '+key)
def run_version(version):
    app=win32com.client.DispatchEx('Excel.Application')
    app.Visible=False; app.DisplayAlerts=False; app.EnableEvents=False
    path=ROOT/'tests'/f'isolated_{version}.xlsm'
    shutil.copy2(BASE['path'],path)
    exports=ROOT/'results'/('run_'+time.strftime('%H%M%S')+'_'+version); exports.mkdir()
    book=None; result={'version':version,'exports':str(exports),'scope':'native Excel, synthetic output ownership + actual small conic solver; not P06N/full25'}
    try:
        book=app.Workbooks.Open(str(path),UpdateLinks=0)
        setting(book,'TEST_INJECTION',1)
        if version=='G1R12':
            for name in CHANGED: replace(book,name,(ROOT/'src'/f'{name}.bas').read_text(encoding='utf-8'))
        else:
            append(book,'RPX_Output','Public Sub RPX_PrepareUpperOutput(ByVal adaptive As Boolean)\nEnd Sub\nPublic Sub RPX_AssertUpperOutput(ByVal adaptive As Boolean)\nEnd Sub')
        append(book,'RPX_Adapt',(ROOT/'tests/AdaptIssueProbe.txt').read_text(encoding='utf-8'))
        for name in ['IssueSolverProbe','MeshProbe','SmallRootProbe']:
            src=ROOT/'tests'/f'{name}.bas'
            if not src.exists():
                old=Path(BASE['path']).parent/'tests'/f'{name}.bas'; shutil.copy2(old,src)
            comp=book.VBProject.VBComponents.Add(1); comp.Name=name
            replace(book,name,src.read_text(encoding='utf-8'))
        cm=book.VBProject.VBComponents('RPX_Solve').CodeModule
        text=cm.Lines(1,cm.CountOfLines)
        marker='    RPX_DiagEnd dpLDL\r\n    If Not factorOK Then'
        assert text.count(marker)==1
        text=text.replace(marker,'    RPX_DiagEnd dpLDL\r\n    If IssueForceFactor Then IssueForceFactor = False: factorOK = False: RPX_PivotFailure = 123\r\n    If Not factorOK Then')
        marker='Public Sub RPX_OptimizeHSD(Optional ByVal prepared As Boolean = False, Optional ByVal auditMode As String = "")'
        pos=text.index(marker)
        tail=text[pos:].replace('    On Error GoTo Failed','    On Error GoTo Failed\r\n    IssueHSDCalls = IssueHSDCalls + 1\r\n    RPX_TestHook "issue_hsd"',1)
        text=text[:pos]+tail
        replace(book,'RPX_Solve',text)
        compile_book(app,book)
        call=lambda name,*args: app.Run("'"+book.Name+"'!"+name,*args)
        result['compile']='PASS'
        result['output']={}
        for variant in ['different_count','same_count','same_owner','retained']:
            got=call('IssueOutput',variant); result['output'][variant]=got; print(version,'output',variant,got,flush=True)
            if version=='G1R12': assert got.startswith('PASS'),got
            elif variant!='same_owner': assert not got.startswith('PASS'),got
        if version=='G1R12':
            result['guards']={}
            for mutation in ['model','field','shape','q','factor','incomplete','epoch']:
                got=call('IssueOutputGuard',mutation); result['guards'][mutation]=got
                print(version,'guard',mutation,got,flush=True); assert got.startswith('ERROR 5:'),got
        result['tokens']=[]
        for n,t,want in json.loads((ROOT/'tests/token_cases.json').read_text(encoding='utf-8')):
            got=bool(call('IssueEligible',n,t,False,False))
            result['tokens'].append({'number':n,'text':t,'expected':want,'actual':got})
            if version=='G1R12': assert got==want,(n,t,want,got)
        if version=='G1R12':
            for cancel,audit in [(True,False),(False,True)]: assert not call('IssueEligible',5,'NEWTON_FACTORIZATION_FAILED pivot=123',cancel,audit)
        result['solver']={}
        for variant in ['normal','factor','memory','subscript','cancel','unknown','audit','model','hsd_failure']:
            got=call('IssueSolve',variant,str(exports/variant))
            result['solver'][variant]=got; print(version,'solver',variant,got,flush=True)
            if version=='G1R12':
                if variant=='normal': assert 'engine=LEGACY;rescues=0;hsd_calls=0;status=OPTIMAL' in got,got
                elif variant=='factor': assert 'engine=HSD_RESCUE;rescues=1;hsd_calls=1;status=OPTIMAL' in got,got
                elif variant in ['memory','subscript','cancel']: assert got.startswith({'memory':'ERROR 7:','subscript':'ERROR 9:','cancel':'ERROR 18:'}[variant]) and 'hsd_calls=0' in got,got
                elif variant in ['unknown','audit','model']: assert got.startswith('ERROR 5:') and 'hsd_calls=0' in got,got
                else: assert got.startswith('ERROR 7:HSD_TEST_MEMORY') and 'hsd_calls=1' in got,got
            elif variant=='factor': assert got.startswith('ERROR 5:NEWTON_FACTORIZATION_FAILED pivot=123') and 'hsd_calls=0' in got,got
        # Actual roots, without injected faults; independently audited later in Python.
        got=call('SmallRoots',str(exports/'roots'))
        result['roots']=got; print(version,'roots',got,flush=True); assert got.startswith('PASS'),got
        result['pass']=True
    finally:
        (ROOT/'results'/f'native_{version}.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
        if book is not None: book.Close(SaveChanges=False)
        app.Quit()
    return result
if __name__=='__main__':
    assert hashlib.sha256(Path(BASE['path']).read_bytes()).hexdigest()==BASE['sha256']
    for version in sys.argv[1:] or ['G1R11','G1R12']: run_version(version)
