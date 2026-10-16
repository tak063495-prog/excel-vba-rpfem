"""Fresh dry-case regression on frozen G1R20 Fix3; no field/root imports."""
from pathlib import Path
import argparse, gc, hashlib, importlib.util, json, shutil, struct, threading, time, uuid
import psutil, pythoncom, win32com.client, win32con, win32gui, win32process

ROOT = Path(__file__).resolve().parents[1]
REPO = ROOT.parents[1]
BASE = REPO / 'versions/G1R19'
FIX = REPO / 'versions/G1R19_2pct'
WATER = REPO / 'versions/G1R19_WaterCommon_20261016'
CANDIDATE = ROOT
BOOK_NAME='RPFEM_20261017_G1R20_Fix3.xlsm'
FROZEN_BOOK_SHA = hashlib.sha256((CANDIDATE/'delivery'/BOOK_NAME).read_bytes()).hexdigest()
FROZEN_SOURCE_MANIFEST = hashlib.sha256((CANDIDATE/'SOURCE_MANIFEST.json').read_bytes()).hexdigest()


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()

def vba(path):
    return '\r\n'.join(s for s in Path(path).read_text(encoding='utf8').splitlines()
                       if not s.startswith('Attribute VB_'))

def sources():
    assert sha(CANDIDATE/'delivery/RPFEM_20261017_G1R20_Fix3.xlsm') == FROZEN_BOOK_SHA
    assert sha(CANDIDATE/'SOURCE_MANIFEST.json') == FROZEN_SOURCE_MANIFEST
    files={p.stem:p for p in (CANDIDATE/'src/vba').glob('*.bas')}
    assert len(files)==63
    for p in files.values():
        assert p.read_text(encoding='utf8')==(CANDIDATE/'import'/p.name).read_text(encoding='cp932')
    return files


def main():
    ap=argparse.ArgumentParser()
    ap.add_argument('--case', choices=['D01A','D02A'], default='D02A')
    ap.add_argument('--guards-only', action='store_true')
    ap.add_argument('--ray-unit', action='store_true')
    ap.add_argument('--precision-from-start', action='store_true')
    ap.add_argument('--budget', type=float, default=900)
    args=ap.parse_args()
    prior=[p.pid for p in psutil.process_iter() if p.name().upper() == 'EXCEL.EXE']
    if args.ray_unit or args.guards_only: raise ValueError('Cold-only entry; micro/guards have separate harness')
    if prior: raise RuntimeError(f'Another Excel exists; refusing: {prior}')
    folder=ROOT/'results'/('owned_excel_'+str(time.time_ns())); folder.mkdir(parents=True)
    bookpath=folder/'Adaptive_Probe.xlsm'
    shutil.copyfile(CANDIDATE/'delivery/RPFEM_20261017_G1R20_Fix3.xlsm', bookpath)
    shutil.copyfile(CANDIDATE/'delivery/rpfem_native_x64.dll', folder/'rpfem_native_x64.dll')
    rec=dict(status='PREPARING',run_token=uuid.uuid4().hex,prior_excel_pids=prior,
             output=str(folder),harness_sha256=sha(__file__),
             source_sha256={n:sha(p) for n,p in sources().items()},
             workbook_event_source_sha256=sha(CANDIDATE/'src/workbook/ThisWorkbook.code.txt'),
             dll_sha256=sha(folder/'rpfem_native_x64.dll'),dialogs=[],
             instrumentation='Test copy adds read-only SaveState checkpoint observer only',
             source_composition='G1R20 Fix3 full63; seven changed modules versus Fix2; ThisWorkbook observer/source persistence added; original mathematical kernels unchanged',
             frozen_candidate_workbook_sha256=FROZEN_BOOK_SHA,source_manifest_sha256=FROZEN_SOURCE_MANIFEST)
    save=lambda:(folder/'summary.json').write_text(json.dumps(rec,indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8')
    save(); shutil.copyfile(__file__,folder/'harness_snapshot.py')
    for name,path in sources().items():
        dest=folder/'source_snapshot'/path.name;dest.parent.mkdir(exist_ok=True);shutil.copyfile(path,dest)
    shutil.copyfile(CANDIDATE/'SOURCE_MANIFEST.json',folder/'SOURCE_MANIFEST.json')
    app=win32com.client.DispatchEx('Excel.Application')
    pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    rec.update(owned_excel_pid=pid,owned_excel_creation_time=psutil.Process(pid).create_time(),owned_excel_cmdline=psutil.Process(pid).cmdline())
    stop=threading.Event(); run_started=[None]
    cancel_stream=pythoncom.CoMarshalInterThreadInterfaceInStream(pythoncom.IID_IDispatch,app._oleobj_)
    def watch():
        pythoncom.CoInitialize()
        cancel_app=win32com.client.Dispatch(pythoncom.CoGetInterfaceAndReleaseStream(cancel_stream,pythoncom.IID_IDispatch))
        while not stop.wait(.2):
            def visit(hwnd,_):
                if win32process.GetWindowThreadProcessId(hwnd)[1]!=pid or not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd)!='#32770':return
                children=[]
                win32gui.EnumChildWindows(hwnd,lambda child,_:children.append((child,win32gui.GetWindowText(child))),None)
                texts=[t for _,t in children]
                if texts not in rec['dialogs']:rec['dialogs'].append(texts)
                for child,label in children:
                    if label in ['OK','&OK'] or ('Visual Basic' in win32gui.GetWindowText(hwnd) and ('終了' in label or '&End' in label)):
                        win32gui.PostMessage(child,win32con.BM_CLICK,0,0)
            try:
                win32gui.EnumWindows(visit,None)
            except Exception as monitor_error:
                code = getattr(monitor_error, "winerror", monitor_error.args[0] if monitor_error.args else None)
                key = "transient_window_1400" if code == 1400 else "window_monitor_errors"
                errors = rec.setdefault(key, [])
                if repr(monitor_error) not in errors:
                    errors.append(repr(monitor_error))
            if run_started[0] is not None:
                elapsed=time.perf_counter()-run_started[0]
                if elapsed > args.budget and not rec.get('cancel_sent'):
                    rec['cancel_requested_elapsed']=elapsed
                    try:
                        cancel_app.Run("'"+bookpath.name+"'!RPX_Stop")
                        rec['cancel_sent']=True
                    except Exception:
                        pass  # Excel may reject the call until its next DoEvents.
                try:
                    proc=psutil.Process(pid)
                    heart=dict(pid=pid,creation_time=rec['owned_excel_creation_time'],elapsed=elapsed,
                               rss=proc.memory_info().rss,cpu=sum(proc.cpu_times()[:2]),cancel_sent=rec.get('cancel_sent',False))
                    (folder/'heartbeat.json').write_text(json.dumps(heart)+'\n',encoding='utf8')
                except psutil.NoSuchProcess:break
        cancel_app=None;pythoncom.CoUninitialize()
    # This watchdog never makes COM calls: a blocked cancellation cannot stall it.
    def hard_watch():
        while not stop.wait(.2):
            if run_started[0] is None:continue
            elapsed=time.perf_counter()-run_started[0]
            if elapsed<=args.budget+45:continue
            try:
                proc=psutil.Process(pid)
                if proc.create_time()!=rec['owned_excel_creation_time'] or proc.cmdline()!=rec['owned_excel_cmdline']:
                    rec['hard_guard_ownership_mismatch']=True;save();return
                rec['hard_terminated_elapsed']=elapsed;save();proc.kill()
            except psutil.NoSuchProcess:pass
            return
    hardwatcher=threading.Thread(target=hard_watch,daemon=True);hardwatcher.start()
    watcher=threading.Thread(target=watch,daemon=True); watcher.start()
    book=component=control=None
    try:
        app.Visible=False; app.DisplayAlerts=False; app.EnableEvents=False
        app.MultiThreadedCalculation.Enabled=False
        book=app.Workbooks.Open(str(bookpath),UpdateLinks=0)
        for name,path in sources().items():
            component=book.VBProject.VBComponents(name)
            component.CodeModule.DeleteLines(1,component.CodeModule.CountOfLines)
            code=vba(path)
            if name=='RPX_Adapt':
                anchor='    state = pending: SaveState = True'
                assert code.count(anchor)==1
                code=code.replace(anchor,anchor+'\r\n    OwnedProductionCheckpoint state, side',1)
            component.CodeModule.AddFromString(code)
        probes={'OwnedProductionProbe':BASE/'tests/OwnedProductionProbe.bas',
                'AdaptiveWeightProbe':FIX/'tests/AdaptiveWeightProbe.bas',
                'LowerRayProbe':FIX/'tests/LowerRayProbe.bas',
                'AccuracyExport':REPO/'validation/G1R12_native25_20261009/scripts/tests/BatchEvidence.bas'}
        for name,path in probes.items():
            component=book.VBProject.VBComponents.Add(1);component.Name=name
            component.CodeModule.AddFromString(vba(path).replace('Private Sub BatchState(', 'Public Sub BatchState('))
        book.Activate();component.CodeModule.CodePane.Show()
        control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled:control.Execute();time.sleep(.3)
        if control.Enabled:
            pane=app.VBE.ActiveCodePane;line=int(pane.GetSelection()[0])
            raise RuntimeError('VBA compile failed '+pane.CodeModule.Name+': '+pane.CodeModule.Lines(max(1,line-2),5))
        app.VBE.MainWindow.Visible=False
        rec['compile']='PASS';save()
        rec['loaded_code_sha256']={c.Name:hashlib.sha256(c.CodeModule.Lines(1,c.CodeModule.CountOfLines).encode('utf8')).hexdigest() for c in book.VBProject.VBComponents if c.Type==1}
        call=lambda name,*values:app.Run("'"+book.Name+"'!"+name,*values)
        rec['weight_unit']=['SKIPPED', 'fresh cold; independent reproduction and geometry micros recorded separately'];save()
        assert rec['weight_unit'][0]=='SKIPPED'
        if args.case:
            path=REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs'/(args.case+'.json')
            inputs=json.loads(path.read_text(encoding='utf8'))
            rec.update(case=args.case,input_sha256=sha(path),imported_field_or_root=False,
                       solver_threads=1,kind='cold_initial_input',budget_seconds=args.budget)
            shutil.copyfile(path,folder/'input_record.json')
            call('RPX_Setup')
            for sheet,ranges in [('要素定義',[('A','C','points',7),('E','H','regions',7),('J','T','faces',7)]),('材料データ',[('A','E','materials',5)])]:
                ws=book.Worksheets(sheet);ws.Range('A7:T1006' if sheet=='要素定義' else 'A5:Q1004').ClearContents()
                for col,last,key,row in ranges:
                    values=tuple(tuple(x) for x in inputs[key]);ws.Range(f'{col}{row}:{last}{row+len(values)-1}').Value2=values
                ws=None
            for key,value in inputs['settings'].items():call('OwnedSetting',key,value)
            call('RPX_G20ImprovementPreset')
            if args.ray_unit:
                rec['lower_ray_unit']=list(call('LowerRayUnit',str(folder/'ray.tsv')));save()
                assert rec['lower_ray_unit'][0]=='PASS',rec['lower_ray_unit']
                rec['direct_homogeneous_unit']=list(call('DirectHomogeneousUnit',str(folder/'direct_homogeneous.tsv')));save()
                assert rec['direct_homogeneous_unit'][0]=='PASS',rec['direct_homogeneous_unit']
                # The unit changes solver settings only; restore the production preset.
                call('RPX_G20ImprovementPreset')
            if not args.guards_only:
                call('OwnedSetting','ADAPT_PRECISION_FROM_START',1) # audited G20 profile, never coarse q4
                call('OwnedSetting','DIAGNOSTIC_MODE',1);call('OwnedSetting','ACCURACY_DLL_CAPTURE',1)
                settings={str(row[2]):row[0] for row in book.Worksheets('設定').Range('B4:D999').Value2 if row[2]}
                rec['settings']=settings
                checkpoints=folder/'checkpoints';checkpoints.mkdir()
                book.Save();rec['configured_workbook_sha256']=sha(bookpath)
                rec.update(status='RUNNING',started_unix=time.time());save()
                print(str(folder)+' cold '+args.case+' start ownedPID='+str(pid),flush=True)
                run_started[0]=time.perf_counter()
                result=list(call('OwnedProductionRun',str(folder/'production.tsv'),str(checkpoints)))
                rec['analysis_seconds']=time.perf_counter()-run_started[0];run_started[0]=None
                rec['cold_result']=result
                rec['result_double_hex']={str(i):struct.pack('<d',v).hex() for i,v in enumerate(result) if isinstance(v,float)}
                rec['target_met']=result[0]=='AUDITED' and bool(result[1])
                rec['status']='PASS' if rec['target_met'] else 'PARTIAL';save()
                if not rec['target_met']:raise RuntimeError('Cold target not achieved: '+repr(result))
                maps=[m.path for m in psutil.Process(pid).memory_maps() if Path(m.path).name.lower().startswith('rpfem_')]
                rec['native_memory_maps']=maps
                assert len(maps)==1 and Path(maps[0]).resolve()==(folder/'rpfem_native_x64.dll').resolve(),maps
                spec=importlib.util.spec_from_file_location('original_replay',BASE.parent/'G1R18/tests/test_accuracy_native_excel.py')
                replay=importlib.util.module_from_spec(spec);spec.loader.exec_module(replay)
                rec['independent_raw_replay']={str(p.relative_to(folder)):replay.replay(p,dual='.accuracy_dual_' in p.name) for p in folder.rglob('*.bin') if '.accuracy_' in p.name}
                book.Save()
        assert sha(CANDIDATE/'SOURCE_MANIFEST.json') == FROZEN_SOURCE_MANIFEST
        assert sha(CANDIDATE/'delivery/RPFEM_20261017_G1R20_Fix3.xlsm') == FROZEN_BOOK_SHA
        assert not rec.get('window_monitor_errors') and not rec['dialogs'], 'Window monitor or dialog failure'
        rec['status']='PASS';save()
        print(json.dumps({'status':'PASS','output':str(folder),'pid':pid,'unit':rec['weight_unit']}),flush=True)
    except BaseException as error:
        rec.update(status='FAIL',error=repr(error));save();raise
    finally:
        try:
            if book is not None:book.Close(False)
        except Exception as cleanup_error:rec['book_cleanup_error']=repr(cleanup_error)
        try:app.Quit()
        except Exception as cleanup_error:rec['app_cleanup_error']=repr(cleanup_error)
        stop.set();watcher.join(2);hardwatcher.join(2)
        book=component=control=app=None;gc.collect()
        for _ in range(50):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        rec['owned_excel_exited']=not psutil.pid_exists(pid);save()
        rec['files']={}
        for p in folder.rglob('*'):
            if not p.is_file() or p.name=='summary.json':continue
            try:rec['files'][str(p.relative_to(folder))]={'bytes':p.stat().st_size,'sha256':sha(p)}
            except PermissionError:rec.setdefault('transient_locked_files',[]).append(str(p.relative_to(folder)))
        save()
        if rec['status']=='PASS' and (not rec['owned_excel_exited'] or rec.get('window_monitor_errors') or rec['dialogs']):
            rec.update(status='FAIL',verification_error='Owned Excel exit, monitor or dialog verification failed')
            save()
            raise RuntimeError(rec['verification_error'])

if __name__=='__main__':main()
