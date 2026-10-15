"""Small serial real-Excel original-LP-first/fallback regression and timings."""
from pathlib import Path
import argparse,gc,hashlib,importlib.util,json,shutil,sys,threading,time,uuid
import psutil,win32com.client,win32con,win32gui,win32process
ROOT=Path(__file__).resolve().parents[1];BASE=ROOT.parent/'G1R18'
sys.path.insert(0,str(BASE/'tools'))
from current_sources import module_paths,sha

def main():
    ap=argparse.ArgumentParser();ap.add_argument('--source-root',type=Path,required=True)
    ap.add_argument('--unified-dll',type=Path);ap.add_argument('--without-accuracy',action='store_true')
    ap.add_argument('--p06-benchmark',action='store_true');ap.add_argument('--terminal',action='store_true');args=ap.parse_args()
    prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
    if prior:raise RuntimeError(f'Serial owned Excel only: {prior}')
    folder=ROOT/'results'/('direct_excel_'+str(time.time_ns()));folder.mkdir(parents=True)
    workbook=folder/'Direct_Test.xlsm';shutil.copyfile(BASE/'delivery/RPFEM_20261014_G1R18.xlsm',workbook)
    if args.unified_dll:shutil.copyfile(args.unified_dll,folder/'rpfem_native_x64.dll')
    else:
        shutil.copyfile(BASE/'delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
        if not args.without_accuracy:shutil.copyfile(BASE/'native/accuracy/bin/rpfem_accuracy_x64.dll',folder/'rpfem_accuracy_x64.dll')
    record=dict(status='PREPARING',output=str(folder),run_token=uuid.uuid4().hex,prior_excel_pids=prior,
                source_root=str(args.source_root.resolve()),harness_sha256=sha(__file__),dll_sha256={p.name:sha(p) for p in folder.glob('*.dll')},cases=[],dialogs=[])
    save=lambda:(folder/'summary.json').write_text(json.dumps(record,indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8')
    shutil.copyfile(__file__,folder/'harness_snapshot.py');save()
    app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    record.update(pid=pid,process_creation_time=psutil.Process(pid).create_time());stop=threading.Event()
    def watch():
        while not stop.wait(.2):
            def visit(hwnd,_):
                if win32process.GetWindowThreadProcessId(hwnd)[1]!=pid or not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd)!='#32770':return
                children=[];win32gui.EnumChildWindows(hwnd,lambda child,_:children.append((child,win32gui.GetWindowText(child))),None)
                texts=[t for _,t in children]
                if texts not in record['dialogs']:record['dialogs'].append(texts)
                for child,text in children:
                    if text in ['OK','&OK'] or ('Visual Basic' in win32gui.GetWindowText(hwnd) and ('終了' in text or '&End' in text)):win32gui.PostMessage(child,win32con.BM_CLICK,0,0)
            win32gui.EnumWindows(visit,None)
    watcher=threading.Thread(target=watch,daemon=True);watcher.start();book=component=module=control=None
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.MultiThreadedCalculation.Enabled=False
        book=app.Workbooks.Open(str(workbook),UpdateLinks=0)
        files=module_paths(args.source_root);record['source_sha256']={name:sha(p) for name,p in files.items()}
        for name,p in files.items():
            module=book.VBProject.VBComponents(name).CodeModule;module.DeleteLines(1,module.CountOfLines)
            module.AddFromString('\r\n'.join(x for x in p.read_text(encoding='utf8').splitlines() if not x.startswith('Attribute VB_')))
        p=Path(__file__).with_name('SpeedDirectProbe.bas');shutil.copyfile(p,folder/p.name)
        component=book.VBProject.VBComponents.Add(1);component.Name='SpeedDirectProbe'
        component.CodeModule.AddFromString('\r\n'.join(x for x in p.read_text(encoding='utf8').splitlines() if not x.startswith('Attribute VB_')))
        if args.terminal:
            p=Path(__file__).with_name('TerminalReuseProbe.bas');shutil.copyfile(p,folder/p.name)
            component=book.VBProject.VBComponents.Add(1);component.Name='TerminalReuseProbe'
            component.CodeModule.AddFromString('\r\n'.join(x for x in p.read_text(encoding='utf8').splitlines() if not x.startswith('Attribute VB_')))
        book.Activate();component.CodeModule.CodePane.Show();control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled:control.Execute();time.sleep(.3)
        if control.Enabled:
            pane=app.VBE.ActiveCodePane
            selected=pane.GetSelection();line=int(selected[0])
            record['compile_error']=dict(module=pane.CodeModule.Name,selection=selected,source=pane.CodeModule.Lines(max(1,line-2),5));save()
            raise RuntimeError('VBA compile failed: '+str(record['compile_error']))
        app.VBE.MainWindow.Visible=False;record['compile']='PASS';save();print(str(folder)+' compile PASS ownedPID='+str(pid),flush=True)
        call=lambda name,*values:app.Run("'"+book.Name+"'!"+name,*values)
        before={name:book.Worksheets(name).UsedRange.Value2 for name in ['要素定義','材料データ']}
        preset=list(call('SpeedPresetProbe'));record['preset']=preset
        assert preset==['PASS',.02,1,8,.0001,8,0],preset
        # Setup creates the baseline sheets before preset; compare the second application.
        before={name:book.Worksheets(name).UsedRange.Value2 for name in before};call('RPX_Speed2pctPreset')
        assert before=={name:book.Worksheets(name).UsedRange.Value2 for name in before}
        for mode,fs in [('lower',.4),('lower',2.),('upper',.4),('upper',2.)]:
            for direct in [0,1]:
                item=list(call('SpeedDirectCase',mode,fs,direct,False,str(folder)));record['cases'].append(item);save()
                assert item[0]=='PASS',item
        item=list(call('SpeedDirectCase','upper',2.,1,True,str(folder)));record['cases'].append(item);save();assert item[0]=='PASS',item
        accepted=[x for x in record['cases'] if x[3]==1 and not x[4] and x[8]]
        if args.without_accuracy:assert not accepted
        else:assert accepted,'Original LP produced no strictly audited direct endpoint'
        if args.p06_benchmark:
            record['p06_fixed_q8_benchmark']=[]
            for case in ['P06A','P06N']:
                path=BASE.parents[1]/'validation/accuracy_20261013/python/data'/(case+'.json')
                inputs=json.loads(path.read_text(encoding='utf8'));shutil.copyfile(path,folder/(case+'_input.json'))
                call('RPX_Setup')
                for sheet,ranges in [('要素定義',[('A','C','points',7),('E','H','regions',7),('J','T','faces',7)]),('材料データ',[('A','E','materials',5)])]:
                    ws=book.Worksheets(sheet);ws.Range('A7:T1006' if sheet=='要素定義' else 'A5:Q1004').ClearContents()
                    for first,last,key,row in ranges:
                        values=tuple(tuple(x) for x in inputs[key]);ws.Range(f'{first}{row}:{last}{row+len(values)-1}').Value2=values
                ws=None
                for key,value in inputs['settings'].items():call('SpeedSetting',key,value)
                exact_inputs={name:book.Worksheets(name).UsedRange.Value2 for name in ['要素定義','材料データ']}
                points=[]
                for flag in [0,1]:
                    item=list(call('SpeedDirectCase','upper',1.,flag,False,str(folder),True));points.append(item)
                    record['p06_fixed_q8_benchmark'].append({'case':case,'input_sha256':sha(path),'result':item,'fixed_fs':1.,'mesh_target':400,'adaptation':False});save()
                    assert item[0]=='PASS',item
                    assert exact_inputs=={name:book.Worksheets(name).UsedRange.Value2 for name in exact_inputs}
                record[case+'_policy']=inputs['settings']['ANALYSIS_POLICY']
        if args.terminal:
            record['terminal']=call('TerminalReuseSmallFixtureRun',str(folder/'terminal.tsv'));save()
            assert str(record['terminal']).startswith('PASS'),record['terminal']
        if args.unified_dll:
            maps=[m.path for m in psutil.Process(pid).memory_maps() if Path(m.path).name.lower().startswith('rpfem_')]
            record['native_memory_maps']=maps
            assert len(maps)==1 and Path(maps[0]).resolve()==(folder/'rpfem_native_x64.dll').resolve(),maps
            assert sha(maps[0])==record['dll_sha256']['rpfem_native_x64.dll'] and len(list(folder.glob('*.dll')))==1
        spec=importlib.util.spec_from_file_location('native_replay',BASE/'tests/test_accuracy_native_excel.py')
        replay=importlib.util.module_from_spec(spec);spec.loader.exec_module(replay)
        record['independent_raw_replay']={str(p.relative_to(folder)):replay.replay(p,dual='.accuracy_dual_' in p.name) for p in folder.rglob('*.bin') if '.accuracy_' in p.name}
        record['status']='PASS';book.Save();save()
        print(json.dumps({'status':'PASS','output':str(folder),'direct_accepts':len(accepted),'cases':record['cases']}),flush=True)
    except BaseException as error:record.update(status='FAIL',error=repr(error));save();raise
    finally:
        try:
            if book is not None:book.Close(SaveChanges=False)
            app.Quit()
        finally:
            stop.set();watcher.join(2);book=component=module=control=app=None;gc.collect()
        for _ in range(50):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        record['owned_excel_exited']=not psutil.pid_exists(pid)
        record['files']={p.relative_to(folder).as_posix():{'sha256':sha(p),'bytes':p.stat().st_size} for p in folder.rglob('*') if p.is_file() and p.name!='summary.json'};save()

if __name__=='__main__':main()
