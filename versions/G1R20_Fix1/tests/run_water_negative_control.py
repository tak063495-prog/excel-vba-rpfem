from pathlib import Path
import gc, hashlib, json, shutil, threading, time, traceback
import psutil, win32com.client, win32con, win32gui, win32process

ROOT=Path(__file__).resolve().parents[1]
REPO=ROOT.parents[1]
SOURCE=REPO/'versions/G1R20/delivery/RPFEM_20261016_G1R20.xlsm'
MACRO=ROOT/'tests/water_negative_control.bas'
OUT=ROOT/'results'/('water_negative_'+str(time.time_ns()))
OUT.mkdir(parents=True)
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
original={str(p):sha(p) for p in (SOURCE,MACRO,REPO/'versions/G1R20/delivery/rpfem_native_x64.dll')}
prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
if prior:raise RuntimeError('Existing Excel; will not take ownership '+str(prior))
bookpath=OUT/'CacheCollisionProbe.xlsm';shutil.copyfile(SOURCE,bookpath)
rec=dict(status='STARTING',solver_calls=0,scope='Public mode1 prescribed producer/cache micro only; no long limit run',
         source_sha256_before=original,python=dict(pid=psutil.Process().pid,create_time=psutil.Process().create_time(),cmdline=psutil.Process().cmdline()),dialogs=[])
def save():(OUT/'summary.json').write_text(json.dumps(rec,indent=2,ensure_ascii=False)+'\n',encoding='utf8')
save();app=book=component=None;stop=threading.Event()
try:
    app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    owner=psutil.Process(pid);birth=owner.create_time();command=owner.cmdline()
    rec['excel']=dict(pid=pid,create_time=birth,cmdline=command,exe=owner.exe());save()
    def watch():
        deadline=time.monotonic()+170
        while not stop.wait(.1):
            if time.monotonic()>deadline:
                if psutil.pid_exists(pid):
                    p=psutil.Process(pid)
                    if p.create_time()==birth and p.cmdline()==command:p.kill();rec['watchdog']='EXACT_OWNED_170S_STOP';save()
                return
            def visit(w,_):
                if win32process.GetWindowThreadProcessId(w)[1]!=pid or win32gui.GetClassName(w)!='#32770':return
                kids=[];win32gui.EnumChildWindows(w,lambda c,_:kids.append((c,win32gui.GetWindowText(c))),None)
                words=[v for _,v in kids]
                if words not in rec['dialogs']:rec['dialogs'].append(words);save()
                for c,v in kids:
                    if v in ('OK','&OK') or '&End' in v or '終了' in v:win32gui.PostMessage(c,win32con.BM_CLICK,0,0)
            win32gui.EnumWindows(visit,None)
    watcher=threading.Thread(target=watch,daemon=True);watcher.start()
    app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.MultiThreadedCalculation.Enabled=False
    book=app.Workbooks.Open(str(bookpath),UpdateLinks=0)
    component=book.VBProject.VBComponents.Add(1);component.Name='WaterCacheCollisionProbe'
    code='\r\n'.join(line for line in MACRO.read_text(encoding='utf8').splitlines() if not line.startswith('Attribute VB_'))
    component.CodeModule.AddFromString(code)
    component.CodeModule.CodePane.Show();control=app.VBE.CommandBars.FindControl(1,578)
    if control.Enabled:control.Execute();time.sleep(.1)
    if control.Enabled:
        pane=app.VBE.ActiveCodePane;line=int(pane.GetSelection()[0])
        raise RuntimeError('Compile failed '+pane.CodeModule.Name+': '+pane.CodeModule.Lines(max(1,line-2),5))
    rec['compile']='PASS';save();app.VBE.MainWindow.Visible=False
    rec['result']=app.Run("'"+book.Name+"'!WaterCacheCollisionRun")
    rec['status']='CONFIRMED_MICRO' if str(rec['result']).startswith('REPRODUCED_STALE_HEAD') else 'NO_CONFIRMATION'
    save()
except BaseException as error:
    rec.update(status='ERROR',error=repr(error),traceback=traceback.format_exc());save()
finally:
    try:
        if book is not None:book.Close(False)
        if app is not None:app.Quit()
    except Exception as error:rec['close_error']=repr(error)
    component=book=app=None;gc.collect()
    stop.set()
    if 'watcher' in globals():watcher.join(2)
    if 'pid' in globals():
        for _ in range(50):
            if not psutil.pid_exists(pid):break
            if psutil.Process(pid).create_time()!=birth:break
            time.sleep(.1)
        rec['owned_excel_exited']=not psutil.pid_exists(pid) or psutil.Process(pid).create_time()!=birth
    rec['source_sha256_after']={p:sha(p) for p in original};rec['production_unchanged']=rec['source_sha256_after']==original
    save()
print(json.dumps(dict(output=str(OUT),status=rec['status'],result=rec.get('result'),excel=rec.get('excel'),
    owned_excel_exited=rec.get('owned_excel_exited'),production_unchanged=rec['production_unchanged']),ensure_ascii=False),flush=True)

raise SystemExit(0 if rec['status']=='CONFIRMED_MICRO' and rec.get('owned_excel_exited') and rec['production_unchanged'] else 1)
