"""Owned Excel independent actual-versus-reported inverse forward evaluation budget regression."""
from pathlib import Path
import argparse, gc, hashlib, importlib.util, json, shutil, threading, time, traceback
import psutil, win32com.client, win32con, win32gui, win32process
ROOT = Path(__file__).resolve().parents[1]
OLD = ROOT.parent / 'G1R20_Fix3'
SPEC = importlib.util.spec_from_file_location('ui_util', OLD.parent/'G1R20_Fix1/tests/UI_run_excel.py')
UTIL = importlib.util.module_from_spec(SPEC); SPEC.loader.exec_module(UTIL)
sha, code, code_key = UTIL.sha, UTIL.code, UTIL.code_key
PROBE = ROOT/'tests/WaterFix4Probe.bas'
OLD_TEMPLATE = OLD/'delivery/RPFEM_20261017_G1R20_Fix3.xlsm'
FINAL_TEMPLATE = ROOT/'delivery/RPFEM_20261017_G1R20_Fix4.xlsm'
TEMPLATE = OLD_TEMPLATE

def main():
    global PROBE, TEMPLATE
    p = argparse.ArgumentParser(); p.add_argument('--execute', action='store_true'); p.add_argument('--negative', action='store_true'); p.add_argument('--output', type=Path, required=True)
    a = p.parse_args()
    TEMPLATE = OLD_TEMPLATE if a.negative else FINAL_TEMPLATE
    if not a.execute: raise SystemExit('Requires coordinated --execute')
    prior = [x.pid for x in psutil.process_iter() if x.name().upper() == 'EXCEL.EXE']
    if prior: raise RuntimeError('Excel slot occupied '+repr(prior))
    out = a.output.resolve(); out.mkdir(parents=True, exist_ok=False)
    sources = OLD if a.negative else ROOT
    paths = sorted((sources/'src/vba').glob('*.bas'))
    if len(paths) != 63: raise ValueError('Expected 63 modules')
    monitored = [TEMPLATE, PROBE, Path(__file__), *paths, *sorted((sources/'import').glob('*.bas'))]
    document = sources/'src/workbook/ThisWorkbook.code.txt'
    if not a.negative:
        if not document.exists(): raise ValueError('Missing Fix3 document module')
        monitored.extend([document, sources/'import/ThisWorkbook.code.txt'])
    if not a.negative: monitored.extend([OLD_TEMPLATE, FINAL_TEMPLATE])
    before = {str(x): sha(x) for x in monitored}
    snap = out/'source_snapshot'; snap.mkdir()
    for path in paths: shutil.copyfile(path, snap/path.name)
    if not a.negative: shutil.copyfile(document, snap/document.name)
    shutil.copyfile(PROBE, out/PROBE.name); shutil.copyfile(__file__, out/'runner_snapshot.py')
    rec = dict(status='PREPARING', negative=a.negative, conic_optimizer_calls=0, product_loading='Existing delivery63+ThisWorkbook verified before adding probe; no product reimport except explicit migration', cases={}, dialogs=[], monitor_errors=[], prior_excel_pids=prior, sha256_before=before, started_epoch=time.time())
    lock = threading.RLock()
    def save():
        with lock: (out/'summary.json').write_text(json.dumps(rec, ensure_ascii=False, indent=2)+'\n', encoding='utf8')
    save(); app = win32com.client.DispatchEx('Excel.Application')
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]; proc = psutil.Process(pid)
    birth, command = proc.create_time(), proc.cmdline()
    rec['excel_owner'] = dict(pid=pid, birth=birth, cmdline=command, exe=proc.exe()); save()
    started = time.monotonic(); stop = threading.Event(); book = component = control = module = None
    def monitor():
        while not stop.wait(.15):
            if time.monotonic()-started > 360:
                try:
                    q = psutil.Process(pid)
                    if q.create_time() == birth and q.cmdline() == command:
                        rec['hard_timeout'] = dict(seconds=360, ownership_matched=True); save(); q.kill()
                except psutil.NoSuchProcess: pass
                return
            def visit(w, _):
                if win32process.GetWindowThreadProcessId(w)[1] != pid or win32gui.GetClassName(w) != '#32770': return
                children=[]; win32gui.EnumChildWindows(w, lambda child,_:children.append((child,win32gui.GetWindowText(child))),None)
                labels=[s for _,s in children]
                if labels not in rec['dialogs']: rec['dialogs'].append(labels); save()
                for child,s in children:
                    if s in ['OK','&OK','終了(&E)','&End']: win32gui.PostMessage(child,win32con.BM_CLICK,0,0)
            try: win32gui.EnumWindows(visit,None)
            except Exception as e:
                item=dict(epoch=time.time(),error=repr(e))
                if getattr(e,'winerror',None)==1400 or (e.args and e.args[0]==1400): rec.setdefault('closed_window_1400',[]).append(item)
                else: rec['monitor_errors'].append(item)
                save()
    watcher=threading.Thread(target=monitor,daemon=True); watcher.start()
    def open_case(label):
        nonlocal book,component,control
        path=out/(label+'.xlsm'); shutil.copyfile(TEMPLATE,path); book=app.Workbooks.Open(str(path),UpdateLinks=0)
        reference = OLD if a.negative or False else ROOT
        for source in sorted((reference/'src/vba').glob('*.bas')):
            component=book.VBProject.VBComponents(source.stem)
            if code_key(component.CodeModule.Lines(1,component.CodeModule.CountOfLines))!=code_key(code(source)):raise AssertionError('Delivery source mismatch '+source.name)
        docref=reference/'src/workbook/ThisWorkbook.code.txt'
        component=book.VBProject.VBComponents('ThisWorkbook')
        if code_key(component.CodeModule.Lines(1,component.CodeModule.CountOfLines))!=code_key(code(docref)):raise AssertionError('Delivery document mismatch')
        component=book.VBProject.VBComponents.Add(1); component.Name=PROBE.stem; component.CodeModule.AddFromString(code(PROBE)); component.CodeModule.CodePane.Show()
        control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled: control.Execute(); time.sleep(.2)
        if control.Enabled: raise RuntimeError('VBA compile failed')
        app.VBE.MainWindow.Visible=False
        loaded=out/(label+'_loaded_source'); loaded.mkdir()
        for x in book.VBProject.VBComponents:
            if x.Type==1 or x.Name=='ThisWorkbook':
                suffix='.code.txt' if x.Name=='ThisWorkbook' else '.bas'
                (loaded/(x.Name+suffix)).write_text(x.CodeModule.Lines(1,x.CodeModule.CountOfLines),encoding='utf8')
    def call(name,*args): return list(app.Run("'"+book.Name+"'!"+name,*args))
    try:
        app.Visible=False; app.DisplayAlerts=False; app.EnableEvents=False; app.MultiThreadedCalculation.Enabled=False
        rec['instrumentation']='Test copy adds one WaterCountForwardEntry call immediately before the existing SKrMin-default branch in RPX_SeepSolve; no equations, data or gates replaced.'
        for maximum in [6, 12, 5]:
            label='inverse_budget_'+str(maximum);open_case(label)
            component=book.VBProject.VBComponents('RPX_Seep');module=component.CodeModule
            original=module.Lines(1,module.CountOfLines);anchor='    If SKrMin <= 0# Then RPX_SeepDefaults'
            if original.count(anchor)!=1:raise AssertionError('Counter instrumentation anchor mismatch')
            changed=original.replace(anchor,'    Call WaterCountForwardEntry\r\n'+anchor,1)
            module.DeleteLines(1,module.CountOfLines);module.AddFromString(changed)
            if code_key(module.Lines(1,module.CountOfLines))!=code_key(changed):raise AssertionError('Counter injection mismatch')
            component.CodeModule.CodePane.Show();control=app.VBE.CommandBars.FindControl(1,578)
            if control.Enabled:control.Execute();time.sleep(.2)
            if control.Enabled:raise RuntimeError('Counter compile failed')
            app.VBE.MainWindow.Visible=False
            (out/(label+'_instrumented_Seep.code.txt')).write_text(module.Lines(1,module.CountOfLines),encoding='utf8')
            result=call('WaterInverseBudget',maximum)
            if maximum>=6:passed=(result[1]==result[2] and result[1]<=maximum and result[4] is True and result[3]<1e-10)
            else:passed=(result[1]==0 and result[4] is not True and 'INVALID_SEEP_INVERSE_MAX' in result[5])
            rec['cases'][label]=dict(result=result,passed=passed);save()
            module=None;book.Save();book.Close(False);book=None;component=control=None
        failures=[k for k,v in rec['cases'].items() if not v['passed']]
        if rec['dialogs'] or rec['monitor_errors']:raise AssertionError('Dialog or monitor failure')
        if a.negative:
            if len(failures)!=len(rec['cases']):raise AssertionError('Expected baseline counter mismatches')
            rec.update(status='CONFIRMED_NEGATIVE',failures=failures)
        else:
            if failures:raise AssertionError('Inverse budget regression failed '+repr(failures))
            rec['status']='PASS'
        save()
    except BaseException as e: rec.update(status='FAIL',error=repr(e),traceback=traceback.format_exc()); save(); raise
    finally:
        try:
            if book is not None: book.Close(False)
        finally:
            try: app.Quit()
            except Exception as e: rec['quit_error']=repr(e)
        module=None;book=component=control=app=proc=None; gc.collect()
        for _ in range(150):
            if not psutil.pid_exists(pid): break
            time.sleep(.1)
        stop.set(); watcher.join(2)
        rec['owned_excel_exited']=not psutil.pid_exists(pid)
        rec['sha256_after']={str(x):sha(x) for x in monitored}
        rec['monitored_files_unchanged']=(rec['sha256_after']==before)
        rec['finished_epoch']=time.time(); rec['elapsed_seconds']=time.monotonic()-started; save()
    if not rec['owned_excel_exited'] or not rec['monitored_files_unchanged'] or rec['monitor_errors'] or rec.get('hard_timeout'):
        rec.update(status='FAIL',error='Ownership, immutability, monitor or timeout failed'); save(); raise AssertionError(rec['error'])
    print(json.dumps(dict(status=rec['status'],cases=len(rec['cases']),output=str(out),owned_excel_exited=rec['owned_excel_exited']),ensure_ascii=False),flush=True)
if __name__=='__main__': main()
