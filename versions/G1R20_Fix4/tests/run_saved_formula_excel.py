"""Saved/reopened formula-input display invalidation; synthetic sentinel, solver0."""
from pathlib import Path
import argparse, gc, importlib.util, json, shutil, threading, time, traceback
import psutil, win32com.client, win32con, win32gui, win32process
ROOT=Path(__file__).resolve().parents[1]; OLD=ROOT.parent/'G1R20_Fix2'
SPEC=importlib.util.spec_from_file_location('saved_ui_util', OLD.parent/'G1R20_Fix1/tests/UI_run_excel.py')
UTIL=importlib.util.module_from_spec(SPEC); SPEC.loader.exec_module(UTIL)
sha,code,code_key=UTIL.sha,UTIL.code,UTIL.code_key
TEMPLATE=OLD/'delivery/RPFEM_20261017_G1R20_Fix2.xlsm'

class Session:
    def __init__(self, rec, save, label):
        self.rec,self.save,self.label=rec,save,label
        self.app=self.book=self.component=self.control=None; self.stop=threading.Event()
    def __enter__(self):
        prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
        if prior: raise RuntimeError('Excel slot occupied '+repr(prior))
        self.app=win32com.client.DispatchEx('Excel.Application')
        self.pid=win32process.GetWindowThreadProcessId(self.app.Hwnd)[1]; proc=psutil.Process(self.pid)
        self.birth,self.command=proc.create_time(),proc.cmdline()
        self.info=dict(label=self.label,pid=self.pid,birth=self.birth,cmdline=self.command,exe=proc.exe(),dialogs=[],monitor_errors=[])
        self.rec['sessions'].append(self.info); self.save(); started=time.monotonic()
        def watch():
            while not self.stop.wait(.15):
                if time.monotonic()-started>120:
                    try:
                        p=psutil.Process(self.pid)
                        if p.create_time()==self.birth and p.cmdline()==self.command:
                            self.info['hard_timeout']=dict(seconds=120,ownership_matched=True);self.save();p.kill()
                    except psutil.NoSuchProcess: pass
                    return
                def visit(w,_):
                    if win32process.GetWindowThreadProcessId(w)[1]!=self.pid or win32gui.GetClassName(w)!='#32770': return
                    children=[];win32gui.EnumChildWindows(w,lambda child,_:children.append((child,win32gui.GetWindowText(child))),None)
                    labels=[s for _,s in children]
                    if labels not in self.info['dialogs']:self.info['dialogs'].append(labels);self.save()
                    for child,s in children:
                        if s in ['OK','&OK','終了(&E)','&End']:win32gui.PostMessage(child,win32con.BM_CLICK,0,0)
                try:win32gui.EnumWindows(visit,None)
                except Exception as e:
                    item=dict(epoch=time.time(),error=repr(e))
                    if getattr(e,'winerror',None)==1400 or (e.args and e.args[0]==1400):self.info.setdefault('closed_window_1400',[]).append(item)
                    else:self.info['monitor_errors'].append(item)
                    self.save()
        self.thread=threading.Thread(target=watch,daemon=True);self.thread.start()
        self.app.Visible=False;self.app.DisplayAlerts=False;self.app.EnableEvents=False;self.app.MultiThreadedCalculation.Enabled=False
        return self
    def call(self,name,*args):return self.app.Run("'"+self.book.Name+"'!"+name,*args)
    def compile(self):
        self.component=self.book.VBProject.VBComponents('RPX_Guard');self.component.CodeModule.CodePane.Show()
        self.control=self.app.VBE.CommandBars.FindControl(1,578)
        if self.control.Enabled:self.control.Execute();time.sleep(.2)
        if self.control.Enabled:raise RuntimeError('VBA compile failed')
        self.app.VBE.MainWindow.Visible=False;self.info['compile']='PASS';self.save()
    def __exit__(self,*_):
        try:
            if self.book is not None:self.book.Close(False)
            if self.app is not None:self.app.Quit()
        except Exception as e:self.info['close_error']=repr(e)
        self.book=self.component=self.control=self.app=None;gc.collect()
        for _ in range(150):
            if not psutil.pid_exists(self.pid):break
            time.sleep(.1)
        self.stop.set();self.thread.join(2);self.info['owned_excel_exited']=not psutil.pid_exists(self.pid);self.save()

def main():
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--negative',action='store_true');p.add_argument('--scenario',choices=['edit','missingproof','openchanged'],default='edit');p.add_argument('--output',type=Path,required=True);a=p.parse_args()
    if not a.execute:raise SystemExit('Requires coordinated --execute')
    if a.negative and a.scenario!='edit':raise ValueError('Negative control uses the edit scenario')
    out=a.output.resolve();out.mkdir(parents=True,exist_ok=False)
    source_root=OLD if a.negative else ROOT;paths=sorted((source_root/'src/vba').glob('*.bas'));doc=source_root/'src/workbook/ThisWorkbook.code.txt'
    if len(paths)!=63:raise ValueError('Expected full63')
    monitored=[TEMPLATE,Path(__file__),*paths,*sorted((source_root/'import').glob('*.bas'))]
    if not a.negative:monitored.extend([doc,source_root/'import/ThisWorkbook.code.txt'])
    before={str(p):sha(p) for p in monitored};snap=out/'source_snapshot';snap.mkdir()
    for p in paths:shutil.copyfile(p,snap/p.name)
    if not a.negative:shutil.copyfile(doc,snap/doc.name)
    shutil.copyfile(__file__,out/'runner_snapshot.py')
    path=out/'SavedFormula.xlsm';shutil.copyfile(TEMPLATE,path)
    rec=dict(status='PREPARING',negative=a.negative,scenario=a.scenario,solver_calls=0,certificate_generated=False,fixture='Synthetic B4:C4 display sentinel; no numerical analysis',sessions=[],sha256_before=before,started_epoch=time.time())
    lock=threading.RLock()
    def save():
        with lock:(out/'summary.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    save();started=time.monotonic()
    try:
        with Session(rec,save,'prepare') as s:
            s.book=s.app.Workbooks.Open(str(path),UpdateLinks=0)
            if not a.negative:
                for name in [x.Name for x in s.book.VBProject.VBComponents if x.Type==1]:s.book.VBProject.VBComponents.Remove(s.book.VBProject.VBComponents(name))
                for source in paths:
                    s.component=s.book.VBProject.VBComponents.Add(1);s.component.Name=source.stem;s.component.CodeModule.AddFromString(code(source))
                    if code_key(s.component.CodeModule.Lines(1,s.component.CodeModule.CountOfLines))!=code_key(code(source)):raise AssertionError('Source mismatch '+source.name)
                s.component=s.book.VBProject.VBComponents('ThisWorkbook')
                if s.component.CodeModule.CountOfLines:s.component.CodeModule.DeleteLines(1,s.component.CodeModule.CountOfLines)
                s.component.CodeModule.AddFromString(code(doc))
            s.compile()
            s.call('RPX_Setup')
            ref=s.book.Worksheets.Add(After=s.book.Worksheets(s.book.Worksheets.Count));ref.Name='参照';ref.Range('A1').Value2=41.65
            materials=s.book.Worksheets('材料データ')
            materials.Range('B5').Formula='=LEN(CELL("filename",A1))' if a.scenario=='openchanged' else "='参照'!A1"
            materials.Calculate()
            result=s.book.Worksheets('解析結果');result.Range('A1').Value2='TEST_SYNTHETIC_SAVED_RESULT';result.Range('B4').Value2=1.11;result.Range('C4').Value2=2.22
            if not a.negative:
                s.call('RPX_InputObserve')
                if a.scenario!='missingproof':s.call('RPX_InputCommit')
            rec['prepared']=dict(material_c=materials.Range('B5').Value2,formula=materials.Range('B5').Formula,sentinel=list(result.Range('B4:C4').Value2[0]),status=result.Range('A1').Value2)
            s.book.Save();rec['saved_sha256']=sha(path);save();ref=materials=result=None
        reopen_path=path
        if a.scenario=='openchanged':
            reopen_path=out/'SavedFormula_reopened_with_a_deliberately_longer_filename.xlsm'
            shutil.copyfile(path,reopen_path)
        rec['reopened_path']=str(reopen_path);save()
        with Session(rec,save,'reopen') as s:
            s.app.EnableEvents=True;s.book=s.app.Workbooks.Open(str(reopen_path),UpdateLinks=0)
            s.compile()
            if not a.negative:
                checks={source.name:code_key(s.book.VBProject.VBComponents(source.stem).CodeModule.Lines(1,s.book.VBProject.VBComponents(source.stem).CodeModule.CountOfLines))==code_key(code(source)) for source in paths}
                checks[doc.name]=code_key(s.book.VBProject.VBComponents('ThisWorkbook').CodeModule.Lines(1,s.book.VBProject.VBComponents('ThisWorkbook').CodeModule.CountOfLines))==code_key(code(doc))
                rec['reopened_source_comparison']=checks
                if not all(checks.values()):raise AssertionError('Reopened product source mismatch')
            result=s.book.Worksheets('解析結果');materials=s.book.Worksheets('材料データ');ref=s.book.Worksheets('参照')
            rec['before_edit']=dict(material_c=materials.Range('B5').Value2,sentinel=list(result.Range('B4:C4').Value2[0]),status=result.Range('A1').Value2)
            if a.scenario=='openchanged':
                rec['open_changed_rejected']=(rec['before_edit']['sentinel']==[None,None] and rec['before_edit']['material_c']!=rec['prepared']['material_c'])
                if not rec['open_changed_rejected']:raise AssertionError('Changed formula at open did not invalidate saved display')
                rec['status']='PASS'
            else:
                rec['same_input_saved_display_retained']=(rec['before_edit']['sentinel']==[1.11,2.22])
                if not rec['same_input_saved_display_retained']:raise AssertionError('Saved sentinel unexpectedly lost before dependency change')
                if a.scenario=='missingproof':
                    rec['proof_missing_labeled']=('入力との対応未確認' in str(rec['before_edit']['status']))
                    if not rec['proof_missing_labeled']:raise AssertionError('Missing proof was not labeled unverified')
                elif not a.negative:
                    rec['matching_proof_preserves_status']=(rec['before_edit']['status']=='TEST_SYNTHETIC_SAVED_RESULT')
                    if not rec['matching_proof_preserves_status']:raise AssertionError('Matching proof unexpectedly changed saved status')
                ref.Range('A1').Value2=51.65;s.app.CalculateFull();materials.Calculate()
                rec['after_edit']=dict(material_c=materials.Range('B5').Value2,formula=materials.Range('B5').Formula,sentinel=list(result.Range('B4:C4').Value2[0]),status=result.Range('A1').Value2)
                if abs(rec['after_edit']['material_c']-51.65)>1e-12:raise AssertionError('Dependency value did not recalculate')
                stale=rec['after_edit']['sentinel']==[1.11,2.22]
                if a.negative:
                    if not stale:raise AssertionError('Negative did not reproduce stale saved display')
                    rec['status']='CONFIRMED_NEGATIVE'
                else:
                    if rec['after_edit']['sentinel']!=[None,None]:raise AssertionError('Recalculated input did not clear saved display')
                    rec['status']='PASS'
            s.book.Save();rec['final_saved_sha256']=sha(reopen_path);save();result=materials=ref=None
    except BaseException as e:rec.update(status='FAIL',error=repr(e),traceback=traceback.format_exc());save();raise
    finally:
        rec['sha256_after']={str(p):sha(p) for p in monitored};rec['monitored_files_unchanged']=(before==rec['sha256_after'])
        rec['finished_epoch']=time.time();rec['elapsed_seconds']=time.monotonic()-started;save()
    if not rec['monitored_files_unchanged'] or any(not s.get('owned_excel_exited') or s['dialogs'] or s['monitor_errors'] or s.get('hard_timeout') for s in rec['sessions']):
        rec.update(status='FAIL',error='Ownership, immutability, dialogs or monitor failed');save();raise AssertionError(rec['error'])
    print(json.dumps(dict(status=rec['status'],output=str(out),sessions=len(rec['sessions']),all_owned_excel_exited=True),ensure_ascii=False),flush=True)
if __name__=='__main__':main()
