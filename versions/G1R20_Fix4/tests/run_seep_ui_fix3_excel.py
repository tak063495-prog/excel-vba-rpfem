"""Owned Excel Fix2-negative/Fix3-positive cancellation and formula regression."""
from pathlib import Path
import argparse, gc, hashlib, importlib.util, json, shutil, threading, time, traceback
import psutil, win32com.client, win32con, win32gui, win32process
ROOT = Path(__file__).resolve().parents[1]
OLD = ROOT.parent / 'G1R20_Fix2'
SPEC = importlib.util.spec_from_file_location('ui_util', OLD.parent/'G1R20_Fix1/tests/UI_run_excel.py')
UTIL = importlib.util.module_from_spec(SPEC); SPEC.loader.exec_module(UTIL)
sha, code, code_key = UTIL.sha, UTIL.code, UTIL.code_key
PROBE = ROOT/'tests/SeepUIFix3Probe.bas'
TEMPLATE = OLD/'delivery/RPFEM_20261017_G1R20_Fix2.xlsm'

def main():
    p = argparse.ArgumentParser(); p.add_argument('--execute', action='store_true'); p.add_argument('--negative', action='store_true'); p.add_argument('--output', type=Path, required=True)
    a = p.parse_args()
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
    before = {str(x): sha(x) for x in monitored}
    snap = out/'source_snapshot'; snap.mkdir()
    for path in paths: shutil.copyfile(path, snap/path.name)
    if not a.negative: shutil.copyfile(document, snap/document.name)
    shutil.copyfile(PROBE, out/PROBE.name); shutil.copyfile(__file__, out/'runner_snapshot.py')
    rec = dict(status='PREPARING', negative=a.negative, conic_optimizer_calls=0, prescribed_field_runs=0, cases={}, dialogs=[], monitor_errors=[], prior_excel_pids=prior, sha256_before=before, started_epoch=time.time())
    lock = threading.RLock()
    def save():
        with lock: (out/'summary.json').write_text(json.dumps(rec, ensure_ascii=False, indent=2)+'\n', encoding='utf8')
    save(); app = win32com.client.DispatchEx('Excel.Application')
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]; proc = psutil.Process(pid)
    birth, command = proc.create_time(), proc.cmdline()
    rec['excel_owner'] = dict(pid=pid, birth=birth, cmdline=command, exe=proc.exe()); save()
    started = time.monotonic(); stop = threading.Event(); book = component = control = None
    def monitor():
        while not stop.wait(.15):
            if time.monotonic()-started > 240:
                try:
                    q = psutil.Process(pid)
                    if q.create_time() == birth and q.cmdline() == command:
                        rec['hard_timeout'] = dict(seconds=240, ownership_matched=True); save(); q.kill()
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
        if not a.negative:
            for name in [x.Name for x in book.VBProject.VBComponents if x.Type==1]: book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(name))
            for source in paths:
                component=book.VBProject.VBComponents.Add(1); component.Name=source.stem; component.CodeModule.AddFromString(code(source))
                if code_key(component.CodeModule.Lines(1,component.CodeModule.CountOfLines))!=code_key(code(source)): raise AssertionError('Source mismatch '+source.name)
            component=book.VBProject.VBComponents('ThisWorkbook')
            if component.CodeModule.CountOfLines: component.CodeModule.DeleteLines(1,component.CodeModule.CountOfLines)
            component.CodeModule.AddFromString(code(document))
            if code_key(component.CodeModule.Lines(1,component.CodeModule.CountOfLines))!=code_key(code(document)): raise AssertionError('Document source mismatch')
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
        for action in ['run','check','preview']:
            label='cancel_'+action; open_case(label); result=call('SeepUIFix3CancelCase',action)
            if action=='run': rec['prescribed_field_runs']+=1
            passed=(result[1]==0 and not result[4] and not result[5] and not result[6] and 'ANALYSIS_CANCELLED' not in result[3])
            if action=='run': passed=passed and result[7] is True and result[3]==''
            if action=='check': passed=passed and '入力に問題は見つかりませんでした。' in result[3]
            if action=='preview': passed=passed and result[3]==''
            rec['cases'][label]=dict(result=result,passed=passed); save(); book.Save(); book.Close(False); book=None; component=control=None
        label='format_formulas'; open_case(label); result=call('SeepUIFix3FormulaCase')
        passed=(result[0]=='' and abs(result[2]-5.165)<1e-12 and not result[5] and result[6]=='')
        rec['cases'][label]=dict(result=result,passed=passed); save(); book.Save(); book.Close(False); book=None; component=control=None
        if not a.negative:
            label='format_legacy_formulas'; open_case(label); result=call('SeepUIFix3LegacyFormulaCase')
            passed=(result[:3]==[3,2,True] and result[3:6]==['=9.81','=16','=1'] and result[6] is True and result[7]==16 and not result[8] and '失敗' not in result[9])
            rec['cases'][label]=dict(result=result,passed=passed); save(); book.Save(); book.Close(False); book=None; component=control=None
            label='busy_format_refused'; open_case(label); result=call('SeepUIFix3BusyFormatCase')
            passed=(result[:4]==[True,True,True,True] and '実行中' in result[4])
            rec['cases'][label]=dict(result=result,passed=passed); save(); book.Save(); book.Close(False); book=None; component=control=None
        failures=[k for k,v in rec['cases'].items() if not v['passed']]
        if rec['dialogs'] or rec['monitor_errors']: raise AssertionError('Dialog or monitor failure')
        if a.negative:
            if set(failures)!=set(rec['cases']): raise AssertionError('Expected every negative case to reproduce '+repr(failures))
            rec.update(status='CONFIRMED_NEGATIVE',failures=failures)
        else:
            if failures: raise AssertionError('Regression failed '+repr(failures))
            rec['status']='PASS'
        save()
    except BaseException as e: rec.update(status='FAIL',error=repr(e),traceback=traceback.format_exc()); save(); raise
    finally:
        try:
            if book is not None: book.Close(False)
        finally:
            try: app.Quit()
            except Exception as e: rec['quit_error']=repr(e)
        book=component=control=app=proc=None; gc.collect()
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
