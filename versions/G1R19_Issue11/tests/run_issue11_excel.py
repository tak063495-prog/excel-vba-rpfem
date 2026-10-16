"""Issue #11 real-Excel probe: G1R19 baseline vs patched RPX_Adapt.

Runs a small q8-from-start fixture (eligible pair) and a q4-coarse fixture
(ineligible) with CYCLES=1 so AfterCycles restores the pair via RPX_RestoreState
(raw phase). Counts final evaluations / terminal reuse events in the trace and
runs the existing TerminalReuseSavedPairProbe mutation checks. Then builds the
patched delivery workbook (only RPX_Adapt replaced), reopens and verifies it.
usage: python run_issue11_excel.py BASE_BOOK SRC_DIR OUT_DIR DELIVERY_NAME
"""
import sys,json,shutil,hashlib,time,csv,gc
from pathlib import Path
import subprocess,win32com.client,win32process
base,src,out,dname=Path(sys.argv[1]),Path(sys.argv[2]),Path(sys.argv[3]),sys.argv[4]
here=Path(__file__).resolve().parent
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def text(p):return '\r\n'.join(x for x in p.read_text(encoding='utf8').splitlines() if not x.startswith('Attribute VB_')).strip()
def compile_book(app,book,module):
    book.Activate();book.VBProject.VBComponents(module).CodeModule.CodePane.Show()
    c=app.VBE.CommandBars.FindControl(1,578)
    if c.Enabled:c.Execute();time.sleep(.4)
    if c.Enabled:raise RuntimeError('compile failed')
    app.VBE.MainWindow.Visible=False
def patch(book):
    cm=book.VBProject.VBComponents('RPX_Adapt').CodeModule;cm.DeleteLines(1,cm.CountOfLines);cm.AddFromString(text(src/'RPX_Adapt.bas'))
def events(trace):
    rows=list(csv.DictReader(trace.open(encoding='utf-16'),delimiter='\t'))  # RPX diagnostic analysis.tsv
    ev=lambda n:[r['detail'] for r in rows if r['event']==n]
    return {'final_evaluations':sum(d.startswith('label=final;') for d in ev('adapt_eval')),
            'terminal_pair_reuse':ev('terminal_pair_reuse'),'terminal_pair_reuse_rejected':len(ev('terminal_pair_reuse_rejected')),
            'adapt_eval':ev('adapt_eval'),'adapt_outcome':ev('adapt_outcome'),'run_end':[r['elapsed_s'] for r in rows if r['event']=='run_end']}
def excel_pids():return [int(l.split('","')[1]) for l in subprocess.run(['tasklist','/FI','IMAGENAME eq EXCEL.EXE','/FO','CSV','/NH'],stdout=subprocess.PIPE).stdout.decode('mbcs','replace').splitlines() if l.upper().startswith('"EXCEL.EXE')]
prior=excel_pids()
if prior:raise SystemExit(f'Another Excel is running {prior}; refusing')
out.mkdir(parents=True,exist_ok=True);shutil.copyfile(base.parent/'rpfem_native_x64.dll',out/'rpfem_native_x64.dll')
app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
res={'excel_version':app.Version,'excel_pid':pid,'base_sha256':sha(base),'patched_source_sha256':sha(src/'RPX_Adapt.bas')};book=None
save=lambda:(out/'issue11_excel.json').write_text(json.dumps(res,indent=2,ensure_ascii=False),encoding='utf8')
try:
    app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
    for variant in ['baseline','patched']:
        res[variant]={}
        for fs in [1,0]:
            tag=f'{variant}-q{8 if fs else 4}';copy=out/f'{tag}.xlsm';shutil.copyfile(base,copy)
            book=app.Workbooks.Open(str(copy),UpdateLinks=0)
            if variant=='patched':patch(book)
            for m in ['Issue11Probe','TerminalReuseProbe']:
                c=book.VBProject.VBComponents.Add(1);c.Name=m;c.CodeModule.AddFromString(text(here/f'{m}.bas'))
            compile_book(app,book,'Issue11Probe')
            run=lambda n,*a:app.Run("'"+book.Name+"'!"+n,*a)
            trace=out/f'{tag}.tsv';trace.unlink(missing_ok=True)
            t0wall=time.time()-1;t=time.perf_counter();r={'result':str(run('Issue11Case',str(trace),fs))};r['seconds']=time.perf_counter()-t
            logs=sorted((p for p in out.rglob('*_analysis.tsv') if p.stat().st_mtime>=t0wall),key=lambda p:p.stat().st_mtime)
            if logs:shutil.copyfile(logs[-1],out/f'{tag}_analysis.tsv');r['diagnostic_log']=logs[-1].name;r['events']=events(logs[-1])
            if fs==1:
                try:r['saved_pair_probe']=str(run('TerminalReuseSavedPairProbe'))
                except Exception as e:r['saved_pair_probe']='ERROR '+repr(e)
            res[variant][f'q8_from_start={fs}']=r;save()
            book.Close(False);book=None
    deliv=out/dname;shutil.copyfile(base,deliv)
    book=app.Workbooks.Open(str(deliv),UpdateLinks=0)
    before={c.Name:(c.CodeModule.Lines(1,c.CodeModule.CountOfLines) if c.CodeModule.CountOfLines else '') for c in book.VBProject.VBComponents}
    patch(book);compile_book(app,book,'RPX_Adapt');book.Save();book.Close(False);book=None
    book=app.Workbooks.Open(str(deliv),UpdateLinks=0,ReadOnly=True);casing=[]
    for c in book.VBProject.VBComponents:
        a=c.CodeModule.Lines(1,c.CodeModule.CountOfLines) if c.CodeModule.CountOfLines else ''
        if c.Name=='RPX_Adapt':
            assert a.strip().lower()==text(src/'RPX_Adapt.bas').lower(),'source mismatch'
        elif a!=before[c.Name]:
            assert a.lower()==before[c.Name].lower(),'unintended change '+c.Name;casing.append(c.Name)
    compile_book(app,book,'RPX_Adapt');book.Close(False);book=None
    res['delivery']={'name':dname,'compile':'PASS','reopen':'PASS','source_match':True,'casing_only':casing,'sha256':sha(deliv)}
    res['status']='DONE'
except Exception as e:
    import traceback;res['status']='FAIL';res['error']=traceback.format_exc()
finally:
    save()
    try:
        if book is not None:book.Close(False)
    except Exception:pass
    try:app.Quit()
    except Exception:pass
    app=None;gc.collect();time.sleep(3)
    if pid in excel_pids():subprocess.run(['taskkill','/PID',str(pid),'/F'])
print(res.get('status'))
