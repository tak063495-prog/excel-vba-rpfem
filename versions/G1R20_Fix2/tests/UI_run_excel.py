"""Owned Excel, solver-free public-preset regression and real G1R19 migration.

Run only after the parent coordinates the single Excel slot. A fresh output
folder records source snapshots and original workbook hashes. No Setup call,
source workbook save, field/root import or native solver call is used.
"""
from pathlib import Path
import argparse,gc,hashlib,json,re,shutil,threading,time,uuid
import psutil,win32com.client,win32con,win32gui,win32process
ROOT=Path(__file__).resolve().parents[1];REPO=ROOT.parents[1]
def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def code(path):return '\r\n'.join(line for line in Path(path).read_text(encoding='utf8').splitlines() if not line.startswith('Attribute VB_'))
def code_key(text):
    lines=[]
    for line in text.replace('\r\n','\n').rstrip('\n').split('\n'):
        line=re.sub(r'^[ \t]*([A-Za-z_]\w*):[ \t]*',r'\1:',line);parts=[];i=0
        while i<len(line):
            if line[i]=="'":parts.append(line[i:]);break
            if line[i]=='"':
                begin=i;i+=1
                while i<len(line):
                    if line[i]=='"':
                        if i+1<len(line) and line[i+1]=='"':i+=2;continue
                        i+=1;break
                    i+=1
                parts.append(line[begin:i]);continue
            parts.append(line[i].lower());i+=1
        lines.append(''.join(parts))
    return hashlib.sha256('\n'.join(lines).encode('utf8')).hexdigest()
def typed(value):
    if isinstance(value,(tuple,list)):return [typed(v) for v in value]
    return [type(value).__name__,value]
def physical(book):
    return {name:{'value2':typed(book.Worksheets(name).UsedRange.Value2),'formula':typed(book.Worksheets(name).UsedRange.Formula)}
            for name in ['要素定義','材料データ','浸透入力']}
def main():
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--output',type=Path,required=True);p.add_argument('--source-root',type=Path,default=ROOT)
    a=p.parse_args()
    if not a.execute:raise SystemExit('Parent must coordinate serial owned Excel; pass --execute explicitly')
    source_root=a.source_root.resolve();paths={f.stem:f for f in (source_root/'src/vba').glob('*.bas')}
    if len(paths)!=63:raise ValueError('A complete frozen 63-module source set is required')
    for path in paths.values():
        if path.read_text(encoding='utf8')!=(source_root/'import'/path.name).read_text(encoding='cp932'):raise ValueError('CP932 pair mismatch '+path.name)
    prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
    if prior:raise RuntimeError('Serial owned Excel required; existing '+repr(prior))
    out=a.output.resolve()
    if out.exists():raise FileExistsError('Refusing result overwrite')
    out.mkdir(parents=True);snap=out/'source_snapshot';snap.mkdir()
    for path in paths.values():shutil.copyfile(path,snap/path.name)
    probe=ROOT/'tests/UI_PresetRegression.bas';shutil.copyfile(probe,out/probe.name);shutil.copyfile(__file__,out/'UI_run_excel_snapshot.py')
    templates={'G20':REPO/'versions/G1R20/delivery/RPFEM_20261016_G1R20.xlsm','G19':REPO/'versions/G1R19/delivery/RPFEM_20261015_G1R19.xlsm'}
    rec={'status':'PREPARING','solver_calls':0,'certificate_generated':False,'setup_called':False,'token':uuid.uuid4().hex,'prior_excel_pids':prior,
         'source_sha256':{n:sha(path) for n,path in paths.items()},'probe_sha256':sha(probe),'harness_sha256':sha(__file__),
         'original_workbook_sha256':{n:sha(path) for n,path in templates.items()},'tests':{},'migration':{},'dialogs':[]}
    save=lambda:(out/'summary.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2,allow_nan=False)+'\n',encoding='utf8')
    save();app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    process=psutil.Process(pid);birth=process.create_time();rec.update(owned_excel_pid=pid,owned_excel_creation_time=birth,owned_command=process.cmdline());save()
    book=c=control=None;stop=threading.Event();started=time.perf_counter()
    def monitor():
        while not stop.wait(.2):
            if time.perf_counter()-started>180:
                if psutil.pid_exists(pid) and psutil.Process(pid).create_time()==birth:
                    rec['hard_timeout']={'seconds':180,'ownership_matched':True};save();psutil.Process(pid).kill()
                return
            def visit(w,_):
                if win32process.GetWindowThreadProcessId(w)[1]!=pid or win32gui.GetClassName(w)!='#32770':return
                kids=[];win32gui.EnumChildWindows(w,lambda c,_:kids.append((c,win32gui.GetWindowText(c))),None)
                texts=[v for _,v in kids]
                if texts not in rec['dialogs']:rec['dialogs'].append(texts);save()
                for child,text in kids:
                    if text in ['OK','&OK','終了(&E)','&End']:win32gui.PostMessage(child,win32con.BM_CLICK,0,0)
            try:win32gui.EnumWindows(visit,None)
            except Exception as error:
                item={'epoch':time.time(),'error':repr(error)}
                if getattr(error,'winerror',None)==1400 or (error.args and error.args[0]==1400):
                    rec.setdefault('monitor_closed_window_1400',[]).append(item)
                else:rec.setdefault('monitor_errors',[]).append(item)
                save()
    thread=threading.Thread(target=monitor,daemon=True);thread.start()
    def open_copy(kind,label):
        nonlocal book,c,control
        path=out/(label+'.xlsm');shutil.copyfile(templates[kind],path);book=app.Workbooks.Open(str(path),UpdateLinks=0)
        before=physical(book)
        for name in [c.Name for c in book.VBProject.VBComponents if c.Type==1]:book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(name))
        for name,path in sorted(paths.items()):
            c=book.VBProject.VBComponents.Add(1);c.Name=name;c.CodeModule.AddFromString(code(path))
            if code_key(c.CodeModule.Lines(1,c.CodeModule.CountOfLines))!=code_key(code(path)):raise ValueError('Source-to-loaded mismatch '+name)
        c=book.VBProject.VBComponents.Add(1);c.Name='UI_PresetRegression';c.CodeModule.AddFromString(code(probe));c.CodeModule.CodePane.Show()
        control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled:control.Execute();time.sleep(.2)
        if control.Enabled:raise RuntimeError('VBA compile failed')
        app.VBE.MainWindow.Visible=False
        if physical(book)!=before:raise ValueError('Source import changed physical workbook inputs')
        return before
    def call(name,*args):return app.Run("'"+book.Name+"'!"+name,*args)
    def settings(keys):return {k:call('RPX_Setting',k,'__MISSING__') for k in keys}
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.MultiThreadedCalculation.Enabled=False
        before=open_copy('G20','Published_G20_UI_TestCopy');policy=call('RPX_Setting','ANALYSIS_POLICY','__MISSING__')
        for entry in range(3):
            for kind in ['protected','late_missing_gap','busy','success']:
                for nested in [False,True]:
                    key=f'entry{entry}_{kind}_nested{int(nested)}';settings_before=typed(book.Worksheets('設定').Range('A4:D1000').Value2)
                    result=call('UI_PresetCase',entry,kind,nested);rec['tests'][key]=result;save()
                    if result[0]!='PASS':raise AssertionError(result)
                    if kind=='busy' and settings_before!=typed(book.Worksheets('設定').Range('A4:D1000').Value2):raise AssertionError('Busy call changed settings')
                    if physical(book)!=before or call('RPX_Setting','ANALYSIS_POLICY','__MISSING__')!=policy:raise AssertionError('Preset changed physical inputs or policy')
        book.Close(False);book=None;c=control=None
        newkeys=['ADAPT_BALANCED_LOWER_TARGET','ADAPT_UPPER_FLOW_PASSES','ACCURACY_DLL_DUAL_REGULARIZATION','ACCURACY_DLL_DUAL_LINEAR_REPAIR','ACCURACY_DLL_UPPER_LP_SECONDS','ACCURACY_DLL_UPPER_RECENTER','FS_TERMINAL_SCHEDULE']
        for entry in range(3):
            before=open_copy('G19',f'Published_G19_Migration_Entry{entry}')
            initial=call('UI_G20Keys');policy=call('RPX_Setting','ANALYSIS_POLICY','__MISSING__')
            if any(item[1] for item in initial):raise ValueError('G19 fixture unexpectedly contains new G20 keys')
            result=call('UI_PresetCase',entry,'migration',False);actual=settings(newkeys)
            expected=dict(zip(newkeys,[0,0,0,0,30,0,0] if entry<2 else [1000,5,1e-8,1,30,1,1]))
            rec['migration'][str(entry)]={'initial_new_keys':initial,'result':result,'settings':actual,'expected':expected,
                'no_setup':True,'complete_source_import_modules':63,'physical_inputs_unchanged':physical(book)==before,'policy_unchanged':call('RPX_Setting','ANALYSIS_POLICY','__MISSING__')==policy};save()
            if result[0]!='PASS' or actual!=expected or not rec['migration'][str(entry)]['physical_inputs_unchanged'] or not rec['migration'][str(entry)]['policy_unchanged']:raise AssertionError(rec['migration'][str(entry)])
            book.Close(False);book=None;c=control=None
        if rec['dialogs']:raise AssertionError('Unexpected modal dialogs')
        if rec.get('monitor_errors'):raise AssertionError('Unexpected monitor error')
        rec['status']='PASS';save()
    except BaseException as error:rec.update(status='FAIL',error=repr(error));save();raise
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
        rec['owned_excel_exited']=not psutil.pid_exists(pid)
        rec['original_workbooks_unchanged']={n:sha(path)==rec['original_workbook_sha256'][n] for n,path in templates.items()}
        rec['source_files_unchanged']={n:sha(path)==rec['source_sha256'][n] for n,path in paths.items()};save()
    if not rec['owned_excel_exited'] or not all(rec['original_workbooks_unchanged'].values()) or not all(rec['source_files_unchanged'].values()) or rec.get('monitor_errors'):
        rec.update(status='FAIL',error='Exit, original workbook or frozen source preservation failed');save()
        raise AssertionError(rec['error'])
    print(json.dumps({'status':rec['status'],'output':str(out),'tests':len(rec['tests']),'real_G19_migrations':len(rec['migration']),'owned_excel_exited':rec['owned_excel_exited']}),flush=True)
if __name__=='__main__':main()
