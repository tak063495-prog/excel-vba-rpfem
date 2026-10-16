"""Serial owned Excel: combined seven-module candidate, zero Fs solves."""
from pathlib import Path
import argparse,gc,hashlib,importlib.util,json,shutil,threading,time,uuid
import psutil,win32com.client,win32con,win32gui,win32process

ROOT=Path(__file__).resolve().parents[1];REPO=ROOT.parents[1]
FIX=REPO/'versions/G1R19_2pct';WATER=REPO/'versions/G1R19_WaterCommon_20261016';BASE=REPO/'versions/G1R19'
spec=importlib.util.spec_from_file_location('owned_harness',FIX/'tests/run_adaptive_excel.py')
h=importlib.util.module_from_spec(spec);spec.loader.exec_module(h)


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--manifest-sha',required=True);args=parser.parse_args()
    assert h.sha(WATER/'SOURCE_MANIFEST.json')==args.manifest_sha
    manifest=json.loads((WATER/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
    files=h.sources()
    for row in manifest['modules']:
        p=WATER/row['source'];imp=WATER/row['imported']
        assert h.sha(p)==row['source_sha256'] and h.sha(imp)==row['import_sha256']
        assert p.read_text(encoding='utf8')==imp.read_text(encoding='cp932')
        files[p.stem]=p
    assert len(files)==62
    prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
    if prior:raise RuntimeError('Another Excel exists: '+repr(prior))
    out=ROOT/'results'/('owned_micro_'+str(time.time_ns()));out.mkdir(parents=True)
    bookpath=out/'WaterCommonMicro.xlsm';shutil.copyfile(BASE/'delivery/RPFEM_20261015_G1R19.xlsm',bookpath)
    shutil.copyfile(BASE/'delivery/rpfem_native_x64.dll',out/'rpfem_native_x64.dll')
    record=dict(status='PREPARING',run_token=uuid.uuid4().hex,source_sha256={n:h.sha(p) for n,p in files.items()},
                water_manifest_sha256=args.manifest_sha,water_manifest=manifest,prior_excel_pids=prior,
                harness_sha256=h.sha(__file__),dll_sha256=h.sha(out/'rpfem_native_x64.dll'),dialogs=[],
                Fs_solver_calls=0,imported_field_root_status=False,scope='Separate seven-module candidate; frozen four-fix artifacts unchanged')
    save=lambda:(out/'summary.json').write_text(json.dumps(record,indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8')
    save();shutil.copyfile(__file__,out/'harness_snapshot.py')
    for n,p in files.items():
        d=out/'source_snapshot'/p.name;d.parent.mkdir(exist_ok=True);shutil.copyfile(p,d)
    probes={'MeshOnlyProbe':FIX/'tests/MeshOnlyProbe.bas','WaterMicroProbe':ROOT/'scripts/WaterMicroProbe.bas'}
    probes.update({p.stem:p for p in (WATER/'tests').glob('WaterCommon*Probe.bas')})
    for p in probes.values():shutil.copyfile(p,out/p.name)
    app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    record.update(owned_excel_pid=pid,owned_excel_creation_time=psutil.Process(pid).create_time());save()
    stop=threading.Event()
    def watch():
        while not stop.wait(.2):
            def visit(w,_):
                if win32process.GetWindowThreadProcessId(w)[1]!=pid or win32gui.GetClassName(w)!='#32770':return
                children=[];win32gui.EnumChildWindows(w,lambda c,_:children.append((c,win32gui.GetWindowText(c))),None)
                words=[x for _,x in children]
                if words not in record['dialogs']:record['dialogs'].append(words)
                for c,t in children:
                    if t in ['OK','&OK'] or ('Visual Basic' in win32gui.GetWindowText(w) and ('終了' in t or '&End' in t)):win32gui.PostMessage(c,win32con.BM_CLICK,0,0)
            win32gui.EnumWindows(visit,None)
    watcher=threading.Thread(target=watch,daemon=True);watcher.start();book=c=control=None
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.MultiThreadedCalculation.Enabled=False
        book=app.Workbooks.Open(str(bookpath),UpdateLinks=0)
        for n,p in files.items():
            c=book.VBProject.VBComponents(n);c.CodeModule.DeleteLines(1,c.CodeModule.CountOfLines);c.CodeModule.AddFromString(h.vba(p))
        for n,p in probes.items():
            c=book.VBProject.VBComponents.Add(1);c.Name=n;c.CodeModule.AddFromString(h.vba(p))
        c.CodeModule.CodePane.Show();control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled:control.Execute();time.sleep(.3)
        if control.Enabled:
            pane=app.VBE.ActiveCodePane;line=int(pane.GetSelection()[0])
            record['compile_error_selection']={'module':pane.CodeModule.Name,'selection':list(pane.GetSelection()),'line':pane.CodeModule.Lines(line,1)};save()
            raise RuntimeError('VBA compile failed '+pane.CodeModule.Name+': '+pane.CodeModule.Lines(max(1,line-2),5))
        app.VBE.MainWindow.Visible=False;record['compile']='PASS';save()
        record['loaded_code_sha256']={c.Name:hashlib.sha256(c.CodeModule.Lines(1,c.CodeModule.CountOfLines).encode('utf8')).hexdigest() for c in book.VBProject.VBComponents if c.Type==1}
        call=lambda name,*v:app.Run("'"+book.Name+"'!"+name,*v)
        call('RPX_Setup');inp=REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs/D08A.json'
        data=json.loads(inp.read_text(encoding='utf8'));shutil.copyfile(inp,out/'input_record.json');record['input_sha256']=h.sha(inp)
        for sheet,ranges in [('要素定義',[('A','C','points',7),('E','H','regions',7),('J','T','faces',7)]),('材料データ',[('A','E','materials',5)])]:
            ws=book.Worksheets(sheet);ws.Range('A7:T1006' if sheet=='要素定義' else 'A5:Q1004').ClearContents()
            for col,last,key,row in ranges:
                vals=tuple(tuple(x) for x in data[key]);ws.Range(f'{col}{row}:{last}{row+len(vals)-1}').Value2=vals
            ws=None
        for k,v in data['settings'].items():call('MeshOnlySetting',k,v)
        call('RPX_Speed2pctPreset')
        record['prescribed_result']=call('WaterMicroPrepareResult');save()
        assert record['prescribed_result'].startswith('PASS'),record['prescribed_result']
        assert call('WaterMicroRollbackSeed')=='PASS'
        record['source_mutations']={}
        for target in ['gamma','gamma_type','coupling','suction','mode','table_x','table_y','result_id','result_head','result_node','result_triangle']:
            result=call('WaterCommonSourceMutationProbe',target);record['source_mutations'][target]=result;save();assert result.startswith('PASS'),result
        call('MeshOnlySetting','TEST_INJECTION',1);record['rollback']={}
        for target in ['restored','arrays','topology','faces','original_cancel','restore_cancel','original_input','original_cancel_restore_failure']:
            assert call('WaterMicroRollbackSeed')=='PASS'
            result=call('WaterCommonRollbackTest',target);record['rollback'][target]=result;save();assert result.startswith('PASS'),result
        call('MeshOnlySetting','TEST_INJECTION',0)
        geometry=out/'geometry';geometry.mkdir()
        record['geometry_probe']=list(call('WaterMicroGenerate',str(geometry)));save()
        assert record['geometry_probe'][0]=='PASS',record['geometry_probe']
        book.Save();record['status']='VBA_MICRO_PASS';save()
    except BaseException as e:record.update(status='FAIL',error=repr(e));save();raise
    finally:
        try:
            if book is not None:book.Close(False)
        finally:app.Quit()
        stop.set();watcher.join(2);book=c=control=app=None;gc.collect()
        for _ in range(50):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        record['owned_excel_exited']=not psutil.pid_exists(pid);save()
    spec=importlib.util.spec_from_file_location('micro_audit',ROOT/'scripts/audit_micro.py')
    audit=importlib.util.module_from_spec(spec);spec.loader.exec_module(audit)
    record['independent_original_geometry_load_audit']=audit.audit(out/'geometry')
    assert h.sha(WATER/'SOURCE_MANIFEST.json')==args.manifest_sha
    record['status']='PASS';record['files']={str(p.relative_to(out)):{'bytes':p.stat().st_size,'sha256':h.sha(p)} for p in out.rglob('*') if p.is_file() and p.name!='summary.json'};save()
    print(json.dumps({'status':record['status'],'output':str(out),'owned_excel_exited':record['owned_excel_exited']}))

if __name__=='__main__':main()
