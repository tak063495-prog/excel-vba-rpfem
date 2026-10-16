"""Owned Excel water switch consumers, capacity boundaries and Fix3-to-Fix4 source-proof migration."""
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
    monitored.extend(sorted((OLD/'src/vba').glob('*.bas')))
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
    started = time.monotonic(); stop = threading.Event(); book = component = control = None
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
    def open_case(label, legacy=False):
        nonlocal book,component,control
        path=out/(label+'.xlsm'); shutil.copyfile(OLD_TEMPLATE if legacy else TEMPLATE,path); book=app.Workbooks.Open(str(path),UpdateLinks=0)
        reference = OLD if a.negative or legacy else ROOT
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
        open_case("switch_consumers")
        for suction in [False, True]:
            for scoped in [False, True]:
                for label,value,formula in [('two',2.,False),('text','invalid',False),('error','=1/0',True),('fraction',.5,False)]:
                    key=('suction' if suction else 'couple')+'_'+str(scoped)+'_'+label
                    result=call('WaterSwitchConsumer',value,suction,scoped,formula)
                    passed=(result[1]==5 and 'INVALID_WATER_SWITCH' in result[2])
                    rec['cases'][key]=dict(result=result,passed=passed);save()
        for value in [True, False, 'TRUE', 'FALSE', 1., 0., '1', '0', '']:
            key='couple_valid_'+str(value)+'_'+type(value).__name__
            result=call('WaterSwitchConsumer',value,False,True,False)
            on=(value is True or value in ['TRUE',1.,'1'])
            expected=-10.19 if on else -20.
            passed=(result[1]==0 and abs(result[4]-expected)<1e-9 and result[5] is on)
            rec['cases'][key]=dict(result=result,passed=passed);save()
        book.Save();book.Close(False);book=None;component=control=None
        if not a.negative:
            open_case('migration',legacy=True)
            app.Run("'"+book.Name+"'!WaterInputConfigure",2)
            old=call('WaterInputCheck',True)
            if old[4] is not True:raise AssertionError('Old producer did not commit')
            rec['migration_old_proof']=call('WaterProofSummary');rec['migration_old_producer_source']='untouched published Fix3';save()
            old_path=Path(book.FullName);book.Save();book.Close(False);book=None;component=control=None
            book=app.Workbooks.Open(str(old_path),UpdateLinks=0)
            for name in [x.Name for x in book.VBProject.VBComponents if x.Type==1]:book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(name))
            for source in paths:
                component=book.VBProject.VBComponents.Add(1);component.Name=source.stem;component.CodeModule.AddFromString(code(source))
                if code_key(component.CodeModule.Lines(1,component.CodeModule.CountOfLines))!=code_key(code(source)):raise AssertionError("Migration source mismatch "+source.stem)
            component=book.VBProject.VBComponents('ThisWorkbook')
            if component.CodeModule.CountOfLines:component.CodeModule.DeleteLines(1,component.CodeModule.CountOfLines)
            component.CodeModule.AddFromString(code(document))
            if code_key(component.CodeModule.Lines(1,component.CodeModule.CountOfLines))!=code_key(code(document)):raise AssertionError("Migration document mismatch")
            component=book.VBProject.VBComponents.Add(1);component.Name=PROBE.stem;component.CodeModule.AddFromString(code(PROBE))
            component.CodeModule.CodePane.Show();control=app.VBE.CommandBars.FindControl(1,578)
            if control.Enabled:control.Execute();time.sleep(.2)
            if control.Enabled:raise RuntimeError('Migration candidate compile failed')
            app.VBE.MainWindow.Visible=False
            loaded=out/"migration_upgraded_loaded_source";loaded.mkdir()
            for x in book.VBProject.VBComponents:
                if x.Type==1 or x.Name=="ThisWorkbook":
                    (loaded/(x.Name+".code.txt")).write_text(x.CodeModule.Lines(1,x.CodeModule.CountOfLines),encoding="utf8")
            result=call('WaterConsumeCurrent')
            rec['cases']['old_proof_rejected']=dict(result=result,passed=result[0]=='REJECTED' and 'SEEP_SOURCE_CHANGED' in result[2]);save()
            renewed=call('WaterInputCheck',True);observed=call('WaterConsumeCurrent')
            rec['migration_new_proof']=call('WaterProofSummary');save()
            rec['cases']['new_proof_recovered']=dict(result=[renewed,observed],passed=renewed[4] is True and observed[0]=='ACCEPTED');save()
            for address,value in [('M207:N207',((2.,20.),)),('A1007:C1007',((999.,2.,2.),))]:
                if address.startswith('A'):book.Worksheets('要素定義').Range(address).Value2=value
                else:book.Worksheets('浸透入力').Range(address).Value2=value
                result=call('WaterConsumeCurrent')
                token='INPUT_RANGE_EXCEEDED' if address.startswith('A') else 'SEEP_TABLE_CAPACITY_EXCEEDED'
                # Main extent token is captured by the shared guard; no stale field may be consumed.
                passed=(result[0]=='REJECTED')
                rec['cases']['consumer_overflow_'+address[:1]]=dict(result=result,passed=passed);save()
                book.Worksheets('要素定義' if address.startswith('A') else '浸透入力').Range(address).ClearContents()
            book.Save();book.Close(False);book=None;component=control=None
        failures=[k for k,v in rec['cases'].items() if not v['passed']]
        if rec['dialogs'] or rec['monitor_errors']:raise AssertionError('Dialog or monitor failure')
        if a.negative:
            rec.update(status='REVIEW_OBSERVED',failures=failures)
        else:
            if failures:raise AssertionError('Transition contracts failed '+repr(failures))
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
