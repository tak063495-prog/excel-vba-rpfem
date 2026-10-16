"""Owned, bounded Excel reproduction/regression of public setup/mesh cleanup.

Each case uses a fresh copy of the published Fix1 workbook and frozen full63
sources (unless --published-only), plus only a test state-machine probe.
No physical witness/root import or numerical solver call occurs.
"""
from pathlib import Path
import argparse,gc,hashlib,importlib.util,json,re,shutil,threading,time,traceback
import psutil,win32com.client,win32con,win32gui,win32process
ROOT=Path(__file__).resolve().parents[1];REPO=ROOT.parents[1]
SPEC=importlib.util.spec_from_file_location('fix1_ui',REPO/'versions/G1R20_Fix1/tests/UI_run_excel.py')
UTIL=importlib.util.module_from_spec(SPEC);SPEC.loader.exec_module(UTIL)
sha=UTIL.sha;code=UTIL.code;code_key=UTIL.code_key;typed=UTIL.typed
def main():
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--published-only',action='store_true');p.add_argument('--output',type=Path,required=True)
    a=p.parse_args()
    if not a.execute:raise SystemExit('Requires coordinated owned Excel: --execute')
    prior=[x.pid for x in psutil.process_iter() if x.name().upper()=='EXCEL.EXE']
    if prior:raise RuntimeError('Excel slot occupied '+repr(prior))
    out=a.output.resolve();out.mkdir(parents=True,exist_ok=False)
    template=REPO/'versions/G1R20_Fix1/delivery/RPFEM_20261017_G1R20_Fix1.xlsm'
    paths={x.stem:x for x in (ROOT/'src/vba').glob('*.bas')}
    if len(paths)!=63:raise ValueError('Expected full63')
    for x in paths.values():
        if x.read_text(encoding='utf8')!=(ROOT/'import'/x.name).read_text(encoding='cp932'):raise ValueError('CP932 mismatch '+x.name)
    snap=out/'source_snapshot';snap.mkdir()
    for x in paths.values():shutil.copyfile(x,snap/x.name)
    probe=ROOT/'tests/UI_SetupControlRegression.bas'
    shutil.copyfile(probe,out/probe.name);shutil.copyfile(__file__,out/'runner_snapshot.py')
    rec={'status':'PREPARING','published_only':a.published_only,'solver_calls':0,'certificate_generated':False,'cases':{},'dialogs':[],
         'prior_excel_pids':prior,'source_sha256':{k:sha(x) for k,x in paths.items()},'probe_sha256':sha(probe),'harness_sha256':sha(__file__),'workbook_event_source_sha256':sha(ROOT/'src/workbook/ThisWorkbook.code.txt'),
         'original_workbook_sha256':sha(template),'started_epoch':time.time()}
    save=lambda:(out/'summary.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    save();app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    proc=psutil.Process(pid);birth=proc.create_time();command=proc.cmdline()
    rec.update(owned_excel_pid=pid,owned_excel_creation_time=birth,owned_command=command);save()
    stop=threading.Event();started=time.monotonic();book=c=control=None
    def monitor():
        while not stop.wait(.2):
            if time.monotonic()-started>240:
                if psutil.pid_exists(pid):
                    q=psutil.Process(pid)
                    if q.create_time()==birth and q.cmdline()==command:
                        rec['hard_timeout']={'seconds':240,'ownership_matched':True};save();q.kill()
                return
            def visit(w,_):
                if win32process.GetWindowThreadProcessId(w)[1]!=pid or win32gui.GetClassName(w)!='#32770':return
                kids=[];win32gui.EnumChildWindows(w,lambda child,_:kids.append((child,win32gui.GetWindowText(child))),None)
                texts=[s for _,s in kids]
                if texts not in rec['dialogs']:rec['dialogs'].append(texts);save()
                for child,s in kids:
                    if s in ['OK','&OK','終了(&E)','&End']:win32gui.PostMessage(child,win32con.BM_CLICK,0,0)
            try:win32gui.EnumWindows(visit,None)
            except Exception as error:
                item={'epoch':time.time(),'error':repr(error)}
                # A VBE/window may disappear after EnumWindows collected it.
                # Only invalid-window 1400 is recoverable; keep watching.
                if getattr(error,'winerror',None)==1400 or (error.args and error.args[0]==1400):
                    rec.setdefault('monitor_closed_window_1400',[]).append(item)
                else:rec.setdefault('monitor_errors',[]).append(item)
                save()
    thread=threading.Thread(target=monitor,daemon=True);thread.start()
    def open_case(label):
        nonlocal book,c,control
        path=out/(label+'.xlsm');shutil.copyfile(template,path);book=app.Workbooks.Open(str(path),UpdateLinks=0)
        if not a.published_only:
            for name in [x.Name for x in book.VBProject.VBComponents if x.Type==1]:book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(name))
            for name,path in sorted(paths.items()):
                c=book.VBProject.VBComponents.Add(1);c.Name=name;c.CodeModule.AddFromString(code(path))
                if code_key(c.CodeModule.Lines(1,c.CodeModule.CountOfLines))!=code_key(code(path)):raise AssertionError('Source mismatch '+name)
        c=book.VBProject.VBComponents('ThisWorkbook');c.CodeModule.DeleteLines(1,c.CodeModule.CountOfLines);c.CodeModule.AddFromString(code(ROOT/'src/workbook/ThisWorkbook.code.txt'))
        c=book.VBProject.VBComponents.Add(1);c.Name='UI_SetupControlRegression';c.CodeModule.AddFromString(code(probe));c.CodeModule.CodePane.Show()
        control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled:control.Execute();time.sleep(.2)
        if control.Enabled:raise RuntimeError('Compile failed')
        app.VBE.MainWindow.Visible=False
    def call(name,*args):return app.Run("'"+book.Name+"'!"+name,*args)
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.MultiThreadedCalculation.Enabled=False
        for kind in ['protected','success','busy']:
            for nested in [False,True]:
                label=f'setup_{kind}_nested{int(nested)}';open_case(label)
                if kind=='busy':book.Worksheets('要素定義').Range('B7').Value2=.125
                before=typed(book.Worksheets('要素定義').UsedRange.Value2)
                result=list(call('UI_SetupCase',kind,nested));after=typed(book.Worksheets('要素定義').UsedRange.Value2)
                expected=([5,'RPX_UI','ANALYSIS_IS_RUNNING',nested,True,True,True] if kind=='busy' else None)
                passed=(result[2]==(5 if kind=='busy' else 1004 if kind=='protected' else 0)
                    and result[5]==nested and result[6]==(kind=='busy')
                    and result[7]==(kind=='busy') and result[8]==(kind=='busy'))
                if kind=='busy':passed=passed and before==after and result[3:5]==['RPX_UI','ANALYSIS_IS_RUNNING']
                rec['cases'][label]={'result':result,'passed':passed,'physical_unchanged':before==after};save()
                book.Close(False);book=None;c=control=None
        for kind in ['prestart_protected_diagram','nested_protected_result','started_protected_nodes','success']:
            open_case('mesh_'+kind)
            # Small geometry-only mesh; no conic or Fs solver.
            book.Worksheets('設定').Range('B5').Value2=16
            result=list(call('UI_MeshCase',kind));isok=kind=='success'
            passed=(result[1]==0 and result[4] is False and result[5] is False and result[6]==isok
                and result[7] is False and result[8]==1 and result[9]==(0 if isok else 1004))
            rec['cases']['mesh_'+kind]={'result':result,'passed':passed};save()
            book.Close(False);book=None;c=control=None
        for entry in range(3):
            for nested in [False,True]:
                label=f'preset_double_entry{entry}_nested{int(nested)}';open_case(label)
                result=list(call('UI_PresetDoubleError',entry,nested))
                passed=(result[2:5]==[5,'RPX_UI','ACCURACY_SETTING_MISSING GAP'] and result[5]==nested
                    and result[6] is False and result[7]==nested and result[8]==nested)
                rec['cases'][label]={'result':result,'passed':passed};save()
                book.Close(False);book=None;c=control=None
        for kind in ['protected_result','invalid_mode']:
            label='run_'+kind;open_case(label);result=list(call('UI_RunCase',kind))
            passed=(result[1]==0 and result[4] is False and result[5] is False
                and result[7] is False and result[8]==1 and result[9]==(1004 if kind=='protected_result' else 5))
            if kind=='invalid_mode':passed=passed and '解析種別はFsまたは極限支持力です。' in result[11]
            rec['cases'][label]={'result':result,'passed':passed};save()
            book.Close(False);book=None;c=control=None
        if rec['dialogs']:raise AssertionError('Unexpected modal dialog')
        if rec.get('monitor_errors'):raise AssertionError('Unexpected monitor error')
        failures=[k for k,v in rec['cases'].items() if not v['passed']]
        if a.published_only:
            if not failures:raise AssertionError('Negative control did not reproduce')
            rec['status']='CONFIRMED_NEGATIVE';rec['failures']=failures
        else:
            if failures:raise AssertionError('Regression failures '+repr(failures))
            rec['status']='PASS'
        save()
    except BaseException as e:rec.update(status='FAIL',error=repr(e),traceback=traceback.format_exc());save();raise
    finally:
        try:
            if book is not None:book.Close(False)
        finally:
            try:app.Quit()
            except Exception:pass
        stop.set();thread.join(2);book=c=control=app=None;gc.collect()
        for _ in range(200):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        rec['owned_excel_exited']=not psutil.pid_exists(pid);rec['original_workbook_unchanged']=sha(template)==rec['original_workbook_sha256']
        rec['source_files_unchanged']={k:sha(x)==rec['source_sha256'][k] for k,x in paths.items()}
        rec['finished_epoch']=time.time();rec['elapsed_seconds']=time.monotonic()-started;save()
    if not rec['owned_excel_exited'] or not rec['original_workbook_unchanged'] or not all(rec['source_files_unchanged'].values()) or rec.get('monitor_errors'):
        rec.update(status='FAIL',error='Ownership, immutability or monitor failed');save();raise AssertionError(rec['error'])
    print(json.dumps({'status':rec['status'],'cases':len(rec['cases']),'failures':rec.get('failures',[]),'output':str(out),'owned_excel_exited':rec['owned_excel_exited']},ensure_ascii=False),flush=True)
if __name__=='__main__':main()
