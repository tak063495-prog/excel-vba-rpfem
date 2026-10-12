import gc,shutil,json
import win32com.client
from prepare import ROOT,write
from native_helpers import compile_book
from test_hsd_regression import setvalue
def run():
    folder=ROOT/'results/issue_regression';folder.mkdir(parents=True,exist_ok=True)
    path=folder/'Issue_Regression.xlsm';shutil.copyfile(ROOT/'delivery/RPFEM_20261012_G1R16_DLL4_Levee_Checked.xlsm',path)
    shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
    app=win32com.client.DispatchEx('Excel.Application');book=None;rows=[]
    def check(label,actual,ok):
        rows.append({'test':label,'actual':actual,'pass':bool(ok)});assert ok,(label,actual)
    try:
        app.Visible=False;app.DisplayAlerts=False;book=app.Workbooks.Open(str(path),UpdateLinks=0)
        setvalue(book,'TEST_INJECTION',1)
        probe=ROOT/'tests/probe'
        book.VBProject.VBComponents('RPX_Adapt').CodeModule.AddFromString((probe/'AdaptIssueProbe.txt').read_text(encoding='utf-8'))
        for name in ['IssueSolverProbe','MeshProbe','SmallRootProbe']:
            c=book.VBProject.VBComponents.Add(1);c.Name=name
            c.CodeModule.AddFromString('\r\n'.join(s for s in (probe/f'{name}.bas').read_text(encoding='utf-8').splitlines() if not s.startswith('Attribute VB_')))
        cm=book.VBProject.VBComponents('RPX_Solve').CodeModule;text=cm.Lines(1,cm.CountOfLines).replace('\r\n','\n')
        marker='    RPX_DiagEnd dpLDL\n    If Not factorOK Then';assert text.count(marker)==1
        text=text.replace(marker,'    RPX_DiagEnd dpLDL\n    If IssueForceFactor Then IssueForceFactor = False: factorOK = False: RPX_PivotFailure = 123\n    If Not factorOK Then')
        start=text.index('Public Sub RPX_OptimizeHSD(');part=text[start:];marker='    On Error GoTo Failed';assert marker in part
        part=part.replace(marker,marker+'\n    IssueHSDCalls = IssueHSDCalls + 1\n    RPX_TestHook "issue_hsd"',1);text=text[:start]+part
        cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(text.replace('\n','\r\n'));compile_book(app,book)
        call=lambda name,*args:app.Run("'"+book.Name+"'!"+name,*args)
        for scenario in ['different_count','same_count','same_owner','retained']:
            actual=str(call('IssueOutput',scenario));check('#1_'+scenario,actual,actual.startswith('PASS'))
        for scenario in ['model','field','shape','q','factor','incomplete','epoch']:
            actual=str(call('IssueOutputGuard',scenario));check('#1_guard_'+scenario,actual,actual.startswith('ERROR 5:'))
        for index,line in enumerate((probe/'tokens.txt').read_text(encoding='utf-8').splitlines()):
            if not line:continue
            number,text,expected=line.split('\t');text=text.replace('\\n','\n').replace('\\r','\r').replace('\\t','\t')
            actual=bool(call('IssueEligible',int(number),text,False,False));check('#2_token_'+str(index),actual,actual==bool(int(expected)))
        for cancel,audit in [(True,False),(False,True)]:
            actual=bool(call('IssueEligible',5,'NEWTON_FACTORIZATION_FAILED pivot=123',cancel,audit));check('#2_cancel_audit',actual,not actual)
        for scenario in ['normal','factor','memory','subscript','cancel','unknown','audit','model','hsd_failure']:
            actual=str(call('IssueSolve',scenario,str(folder/('solver_'+scenario))))
            if scenario=='normal':ok='engine=LEGACY;rescues=0;hsd_calls=0;status=OPTIMAL' in actual
            elif scenario=='factor':ok='engine=HSD_RESCUE;rescues=1;hsd_calls=1;status=OPTIMAL' in actual
            elif scenario=='hsd_failure':ok=actual.startswith('ERROR 7:HSD_TEST_MEMORY') and 'hsd_calls=1' in actual
            else:ok=actual.startswith('ERROR '+str({'memory':7,'subscript':9,'cancel':18}.get(scenario,5))+':') and 'hsd_calls=0' in actual
            check('#2_solver_'+scenario,actual,ok);print(scenario,actual,flush=True)
        actual=str(call('SmallRoots',str(folder/'c5_roots')));check('C5_roots',actual,actual.startswith('PASS'));print(actual,flush=True)
    finally:
        write(folder/'summary.json',{'tests':rows,'passed':sum(r['pass'] for r in rows),'failed':sum(not r['pass'] for r in rows),'scope':'same G1R16 issue #1/#2 assertions; test-only injection in isolated workbook'})
        if book is not None:book.Close(False)
        app.Quit();book=None;app=None;gc.collect()
    print('ISSUE regression PASS',len(rows),flush=True)
if __name__=='__main__':run()
