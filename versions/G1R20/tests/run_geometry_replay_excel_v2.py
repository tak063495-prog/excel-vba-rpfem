"""Serial owned Excel: G20 geometry-only entry and a bounded fresh fixed-Fs point.

This never opens an archived field/root/certificate. A chosen fixed Fs is an
explicit comparison argument, not an imported bound. It does not claim a root.
"""
from pathlib import Path
import argparse,gc,hashlib,importlib.util,json,shutil,sys,threading,time,uuid
import psutil,win32com.client,win32con,win32gui,win32process

ROOT=Path(__file__).resolve().parents[1];REPO=ROOT.parents[1]
VALID=REPO/'validation/G1R20_native_geometry_replay_20261016'
BASE=REPO/'versions/G1R19/delivery'
DLL_SHA='66da1a21be43cd2b3a3bfbaa9d13aad5fb1fd1960e1bf2507cc4e1cfa9ebe70b'
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def code(p):return '\r\n'.join(s for s in Path(p).read_text(encoding='utf8').splitlines() if not s.startswith('Attribute VB_'))
def module(name,path):
    spec=importlib.util.spec_from_file_location(name,path);m=importlib.util.module_from_spec(spec);spec.loader.exec_module(m);return m

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--mode',choices=['micro','fixed'],default='micro')
    parser.add_argument('--case',choices=['P06A','P06N'],default='P06A');parser.add_argument('--side',choices=['lower','upper'],default='upper')
    parser.add_argument('--fs',type=float);parser.add_argument('--hard-budget',type=float,default=600)
    args=parser.parse_args()
    if args.mode=='fixed' and (args.fs is None or not 0<args.fs<100):raise ValueError('An explicit finite comparison Fs is required')
    prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
    if prior:raise RuntimeError('Serial Excel required '+repr(prior))
    out=VALID/'results'/('owned_'+args.mode+'_'+str(time.time_ns()));out.mkdir(parents=True)
    source_manifest=ROOT/'SOURCE_MANIFEST.json';manifest=json.loads(source_manifest.read_text(encoding='utf8'))
    paths={p.stem:p for p in (ROOT/'src/vba').glob('*.bas')};assert len(paths)==63
    for name,item in manifest['modules'].items():
        p=ROOT/'src/vba'/name;assert sha(p)==item['utf8_sha256'],name
    bookpath=out/'G20GeometryReplay.xlsm';shutil.copyfile(BASE/'RPFEM_20261015_G1R19.xlsm',bookpath)
    shutil.copyfile(BASE/'rpfem_native_x64.dll',out/'rpfem_native_x64.dll');assert sha(out/'rpfem_native_x64.dll')==DLL_SHA
    source_out=out/'source_snapshot';source_out.mkdir()
    for p in paths.values():shutil.copyfile(p,source_out/p.name)
    tests=[ROOT/'tests/GeometryReplayProbe.bas',ROOT/'tests/G20PrivateProbe.bas.txt',ROOT/'tests/TerminalRequestProbe.bas.txt',ROOT/'tests/G20PresetErrorProbe.bas']
    for p in tests:shutil.copyfile(p,out/p.name)
    shutil.copyfile(source_manifest,out/'SOURCE_MANIFEST.json');shutil.copyfile(__file__,out/'harness_snapshot.py')
    record=dict(status='PREPARING',mode=args.mode,case=args.case,side=args.side,comparison_fs=args.fs,
        scope='Geometry-only replay then fresh original native fixed point; no completed root claimed',
        archived_primal_imported=False,archived_root_imported=False,token=uuid.uuid4().hex,prior_excel_pids=prior,
        source_manifest_sha256=sha(source_manifest),source_sha256={n:sha(p) for n,p in paths.items()},
        harness_sha256=sha(__file__),dll_sha256=DLL_SHA,dialogs=[],results={})
    save=lambda:(out/'summary.json').write_text(json.dumps(record,indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8')
    save();app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    process=psutil.Process(pid);creation=process.create_time();record.update(owned_excel_pid=pid,owned_excel_creation_time=creation,owned_command=process.cmdline());save()
    stop=threading.Event();started_all=time.perf_counter()
    def monitor():
        while not stop.wait(.2):
            if time.perf_counter()-started_all>args.hard_budget:
                if psutil.pid_exists(pid) and abs(psutil.Process(pid).create_time()-creation)<.01 and psutil.Process(pid).cmdline()==record['owned_command']:
                    record['hard_timeout']={'budget_seconds':args.hard_budget,'ownership_matched':True};save();psutil.Process(pid).kill()
                return
            def visit(w,_):
                if win32process.GetWindowThreadProcessId(w)[1]!=pid or win32gui.GetClassName(w)!='#32770':return
                kids=[];win32gui.EnumChildWindows(w,lambda c,_:kids.append((c,win32gui.GetWindowText(c))),None)
                texts=[t for _,t in kids]
                if texts not in record['dialogs']:record['dialogs'].append(texts)
                for c,t in kids:
                    if t in ['OK','&OK'] or ('Visual Basic' in win32gui.GetWindowText(w) and ('終了' in t or '&End' in t)):win32gui.PostMessage(c,win32con.BM_CLICK,0,0)
            win32gui.EnumWindows(visit,None)
    watcher=threading.Thread(target=monitor,daemon=True);watcher.start();book=c=control=ws=None
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.MultiThreadedCalculation.Enabled=False
        book=app.Workbooks.Open(str(bookpath),UpdateLinks=0)
        for n in [c.Name for c in book.VBProject.VBComponents if c.Type==1]:book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(n))
        for n,p in sorted(paths.items()):
            c=book.VBProject.VBComponents.Add(1);c.Name=n;c.CodeModule.AddFromString(code(p))
        # Private-access bridges exist only in this owned test copy.
        for name,p in [('RPX_Adapt',ROOT/'tests/G20PrivateProbe.bas.txt'),('RPX_AccuracyNative',ROOT/'tests/TerminalRequestProbe.bas.txt')]:
            c=book.VBProject.VBComponents(name);c.CodeModule.AddFromString(code(p))
        upper=REPO/'versions/G1R19_UpperRecovery_20261016/tests/UpperRecoveryArrayProbe.bas'
        c=book.VBProject.VBComponents('RPX_AccuracyNative');c.CodeModule.AddFromString(code(upper).replace('Option Explicit',''))
        c=book.VBProject.VBComponents.Add(1);c.Name='GeometryReplayProbe';c.CodeModule.AddFromString(code(ROOT/'tests/GeometryReplayProbe.bas'))
        c=book.VBProject.VBComponents.Add(1);c.Name='G20PresetErrorProbe';c.CodeModule.AddFromString(code(ROOT/'tests/G20PresetErrorProbe.bas'))
        batch=code(REPO/'validation/G1R12_native25_20261009/scripts/tests/BatchEvidence.bas').replace('Private Sub BatchState','Public Sub BatchState')
        batch=batch.replace('    If Not state.complete Then err.Raise 5, "BatchEvidence", "STATE_INCOMPLETE"','    \' Fixed-point test export only; completeness is separately false in result metadata.')
        c=book.VBProject.VBComponents.Add(1);c.Name='AccuracyExport';c.CodeModule.AddFromString(batch)
        c.CodeModule.CodePane.Show();control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled:control.Execute();time.sleep(.3)
        if control.Enabled:
            pane=app.VBE.ActiveCodePane;line=int(pane.GetSelection()[0]);record['compile_error']={'module':pane.CodeModule.Name,'line':line,'code':pane.CodeModule.Lines(max(1,line-2),5)};save();raise RuntimeError('VBA compile failed')
        app.VBE.MainWindow.Visible=False;record['compile']='PASS';save()
        call=lambda n,*v:app.Run("'"+book.Name+"'!"+n,*v)
        call('RPX_Setup');input_path=REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs'/(args.case+'.json')
        data=json.loads(input_path.read_text(encoding='utf8'));record['canonical_input_sha256']=sha(input_path);shutil.copyfile(input_path,out/'canonical_input.json')
        for sheet,ranges in [('要素定義',[('A','C','points',7),('E','H','regions',7),('J','T','faces',7)]),('材料データ',[('A','E','materials',5)])]:
            ws=book.Worksheets(sheet);ws.Range('A7:T1006' if sheet=='要素定義' else 'A5:Q1004').ClearContents()
            for first,last,key,row in ranges:
                vals=tuple(tuple(v) for v in data[key]);ws.Range(f'{first}{row}:{last}{row+len(vals)-1}').Value2=vals
        ws=book.Worksheets('設定')
        for key,value in data['settings'].items():
            rows=ws.Range('B4:D1000').Value2;chosen=next((i for i,r in enumerate(rows,4) if r[2]==key),None)
            if chosen is None:chosen=next(i for i,r in enumerate(rows,4) if r[2] in [None,''] and r[0] in [None,'']);ws.Cells(chosen,4).Value2=key
            ws.Cells(chosen,2).Value2=value
        ws=None;call('RPX_G20ImprovementPreset')
        keys=['GAP','QUADRATURE','FS_TOL','DIRECT_DLL','ADAPT_PRECISION_FROM_START','ADAPT_C0_WEDGE_AUTO','ADAPT_BALANCED_LOWER_TARGET','ADAPT_UPPER_FLOW_PASSES','FS_TERMINAL_SCHEDULE','ACCURACY_DLL_UPPER_RECENTER','ACCURACY_DLL_DUAL_LINEAR_REPAIR','ACCURACY_DLL_DUAL_REGULARIZATION','DIRECT_DLL_SECONDS','ACCURACY_DLL_DUAL_SECONDS','ACCURACY_DLL_DUAL_EQUILIBRATE','ACCURACY_DLL_SECONDS','ACCURACY_DLL_EQUILIBRATE']
        record['settings']={k:call('RPX_Setting',k,-1) for k in keys};save()
        binary=VALID/'geometry'/(args.case+'_'+args.side+'.bin');record['geometry_binary_sha256']=sha(binary)
        record['results']['geometry_load']=call('GeometryReplayLoadUnit',str(binary));save();assert record['results']['geometry_load'][0]=='PASS'
        if args.mode=='micro':
            record['results']['preset_error_cleanup']=call('G20PresetErrorCleanupUnit');save();assert record['results']['preset_error_cleanup'][0]=='PASS'
            record['results']['upper_work_array']=call('RPX_TestUpperWorkArrays');save();assert record['results']['upper_work_array'].startswith('PASS')
            record['results']['terminal_request']=call('G20TerminalRequestUnit');save();assert record['results']['terminal_request'][0]=='PASS'
            record['results']['balanced_red4']=call('G20BalancedGeometryUnit',1000,str(out/'balanced_geometry.bin'));save();assert record['results']['balanced_red4'][0]=='PASS'
        else:
            t=time.perf_counter();record['results']['fixed_point']=call('GeometryReplayFixedPoint',str(binary),args.side,args.fs,str(out/'fixed'))
            record['fixed_point_seconds']=time.perf_counter()-t;save()
            assert record['results']['fixed_point'][0]=='AUDITED_POINT',record['results']['fixed_point']
        record['native_memory_maps']=[m.path for m in process.memory_maps() if Path(m.path).name.lower().startswith('rpfem_')]
        if not record['native_memory_maps']:assert call('RPX_NativeAvailable');record['native_memory_maps']=[m.path for m in process.memory_maps() if Path(m.path).name.lower().startswith('rpfem_')]
        assert record['native_memory_maps']==[str(out/'rpfem_native_x64.dll')]
        book.Save();record['status']='PASS';save()
    except BaseException as e:record.update(status='FAIL',error=repr(e));save();raise
    finally:
        try:
            if book is not None:book.Close(False)
        finally:
            try:app.Quit()
            except Exception:pass
        stop.set();watcher.join(2);book=c=control=ws=app=None;gc.collect()
        for _ in range(200):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        record['owned_excel_exited']=not psutil.pid_exists(pid)
        record['files']={str(p.relative_to(out)):{'bytes':p.stat().st_size,'sha256':sha(p)} for p in out.rglob('*') if p.is_file() and p.name!='summary.json'};save()
    print(json.dumps({'status':record['status'],'output':str(out),'results':record['results'],'owned_excel_exited':record['owned_excel_exited']}),flush=True)
if __name__=='__main__':main()
