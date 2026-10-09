from pathlib import Path
import sys,json,shutil,time,gc,hashlib,traceback
import pythoncom,win32com.client,win32process
ROOT=Path(__file__).resolve().parents[1]
OLD=Path('C:/Users/link_/Desktop/RPFEM_G1R12_Native25_20261009')
DIAG=Path('C:/Users/link_/Desktop/RPFEM_D03N_Trace_20261009')
REPO=Path('C:/Users/link_/Downloads/excel-vba-rpfem_G1R11_20261008')
sys.path[:0]=[str(OLD),str(REPO/'versions/G1R12/tests')]
from native_helpers import compile_book
from case_locations import case_folder
from check_saved_sources import normalize
from datetime import datetime,timezone
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def write(p,obj):
    p.parent.mkdir(parents=True,exist_ok=True);t=p.with_suffix('.tmp');t.write_text(json.dumps(obj,ensure_ascii=False,indent=2),encoding='utf-8');t.replace(p)
def call(app,book,name,*args):return app.Run("'"+book.Name+"'!"+name,*args)
def inject(book,name,text):
    try:c=book.VBProject.VBComponents(name);c.CodeModule.DeleteLines(1,c.CodeModule.CountOfLines)
    except Exception:c=book.VBProject.VBComponents.Add(1);c.Name=name
    c.CodeModule.AddFromString('\r\n'.join(row for row in text.splitlines() if not row.startswith('Attribute VB_')))
def prepare(case,point=False):
    folder=ROOT/('single_point' if point else 'cases')/case;folder.mkdir(parents=True,exist_ok=True)
    original=case_folder(case)/f'G1R12_{case}_Full1536_Q8.xlsm'
    bookpath=folder/f'RPFEM_20261010_G1R13_{case}.xlsm';assert not bookpath.exists();shutil.copy2(original,bookpath)
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.EnableEvents=False;app.DisplayAlerts=False;book=None
    try:
        book=app.Workbooks.Open(str(bookpath),UpdateLinks=0)
        for p in sorted((ROOT/'src/vba').glob('*.bas')):inject(book,p.stem,p.read_text(encoding='utf-8'))
        if point:
            replay=(DIAG/'src/D03Replay.bas').read_text(encoding='utf-8').replace('    D03TraceRawProblem rawPath\n','')
            inject(book,'D03Replay',replay)
            inject(book,'RepairProbe',(ROOT/'tests/RepairProbe.bas').read_text(encoding='utf-8'))
        compile_book(app,book)
        precision=call(app,book,'RPX_TestHSDPrecision');assert precision.startswith('PASS '),precision
        matched=0
        for p in sorted((ROOT/'src/vba').glob('*.bas')):
            cm=book.VBProject.VBComponents(p.stem).CodeModule
            assert normalize(cm.Lines(1,cm.CountOfLines))==normalize(p.read_text(encoding='utf-8')),p.name;matched+=1
        definition=call(app,book,'BatchDefinition').split('\t')
        book.Save();write(folder/'prepare.json',{'compile':'PASS','precision':precision,'source_module_matches':matched,'definition':definition,'original_sha256':sha(original),'book_sha256':sha(bookpath)})
        print('PREPARED',case,'POINT' if point else 'FULL',precision,flush=True)
    finally:
        if book is not None:book.Close(SaveChanges=False)
        app.Quit();gc.collect()
def point_run():
    folder=ROOT/'single_point/D03N';bookpath=folder/'RPFEM_20261010_G1R13_D03N.xlsm';app=book=None
    rec={'status':'STARTING','scope':'targeted actual failed mesh and exact final Fs; preparation uses reconstructed intermediate factors'};start=time.perf_counter()
    try:
        app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
        rec['excel_pid']=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        book=app.Workbooks.Open(str(bookpath),UpdateLinks=0)
        call(app,book,'BatchStagePath',str(folder/'stage.txt'));call(app,book,'RPX_DiagStart','repair_point')
        seq=tuple(r['value'] for r in json.loads((DIAG/'single_point/factor_reconstruction.json').read_text(encoding='utf-8'))['selected'])
        rec['prepare']=call(app,book,'D03ReplayPrepare',str(DIAG/'capture/candidate_s59_c1_i50.mesh.bin'),str(folder/'raw_problem.bin'),seq,str(DIAG/'single_point/model_aux.csv'),str(DIAG/'single_point/expected_model.txt'))
        assert rec['prepare'].startswith('READY '),rec
        rec['control']=call(app,book,'RepairOriginalControl',str(DIAG/'capture/candidate_s59_c1_i50.mesh.bin'),str(folder/'original_normalized.x.bin')).split('\t')
        assert rec['control'][0]=='PASS',rec['control']
        rec['status']='RUNNING';write(folder/'status.json',rec);print('START POINT',rec['excel_pid'],flush=True)
        rec['native_result']=call(app,book,'D03ReplayRun').split('\t')
        rec['status']='PASS' if rec['native_result'][0]=='0' and rec['native_result'][1]=='OPTIMAL' else 'FAIL'
        if rec['status']=='PASS':rec['final_audit']=call(app,book,'RepairFinalAudit',str(folder/'solved.x.bin')).split('\t')
        book.Save()
    except Exception:rec.update(status='HARNESS_ERROR',traceback=traceback.format_exc())
    finally:
        rec.update(seconds=time.perf_counter()-start,finished_utc=datetime.now(timezone.utc).isoformat());write(folder/'status.json',rec)
        if book is not None:book.Close(SaveChanges=False)
        if app is not None:app.Quit()
        gc.collect()
    print('DONE POINT',json.dumps(rec,ensure_ascii=False),flush=True)
if __name__=='__main__':
    if sys.argv[1]=='prepare':prepare(sys.argv[2],len(sys.argv)>3 and sys.argv[3]=='point')
    elif sys.argv[1]=='point':point_run()
