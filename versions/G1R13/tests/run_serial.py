import os
for key in ['OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS']:os.environ[key]='1'
from pathlib import Path
import json,csv,shutil,hashlib,time,sys,gc,traceback,concurrent.futures,multiprocessing
from datetime import datetime,timezone
import win32com.client,win32process,pythoncom
sys.path.append('C:/Users/link_/Desktop/RPFEM_G1R12_Native25_20261009')
from prepare_inputs import REPO,SOURCE,SOURCE_SHA
ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(REPO/'versions/G1R12/tests'))
from native_helpers import compile_book
from check_saved_sources import normalize
from grade_results import grade
TEMPLATE=ROOT/'cases/D07A/RPFEM_20261010_G1R13_D07A.xlsm'
def utc():return datetime.now(timezone.utc).isoformat()
def write(path,data):
    temp=path.with_suffix(path.suffix+'.tmp');temp.write_text(json.dumps(data,ensure_ascii=False,indent=2),encoding='utf-8',newline='\n');temp.replace(path)
def setvalue(book,key,value):
    ws=book.Worksheets('設定')
    for row in range(4,101):
        if ws.Cells(row,4).Value2==key:ws.Cells(row,2).Value2=value;return
    raise ValueError('SETTING_NOT_FOUND '+key)
def prepare_template():
    assert not TEMPLATE.exists() and hashlib.sha256(SOURCE.read_bytes()).hexdigest()==SOURCE_SHA
    shutil.copy2(SOURCE,TEMPLATE)
    app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;book=None
    try:
        book=app.Workbooks.Open(str(TEMPLATE),UpdateLinks=0)
        comp=book.VBProject.VBComponents.Add(1);comp.Name='BatchEvidence'
        comp.CodeModule.AddFromString('\r\n'.join((ROOT/'tests/BatchEvidence.bas').read_text(encoding='utf-8').splitlines()[1:]))
        compile_book(app,book)
        for p in (REPO/'versions/G1R12/src').glob('*.bas'):
            cm=book.VBProject.VBComponents(p.stem).CodeModule
            assert normalize(cm.Lines(1,cm.CountOfLines))==normalize(p.read_text(encoding='utf-8')),p.name
        book.Save()
        write(ROOT/'results/template_check.json',{'compile':'PASS','production_modules_unchanged':46,'only_added_module':'BatchEvidence','scope':'read-only evidence export; no solver modification','sha256':hashlib.sha256(TEMPLATE.read_bytes()).hexdigest()})
    finally:
        if book is not None:book.Close(SaveChanges=False)
        app.Quit()
def worker(case):
    pythoncom.CoInitialize()
    app=None;book=None;started=time.perf_counter();folder=ROOT/'native_retry'/case;folder.mkdir(parents=True,exist_ok=True)
    rec=json.loads((ROOT/'inputs'/f'{case}.json').read_text(encoding='utf-8'))
    path=folder/f'RPFEM_20261010_G1R13_{case}_Full1536_Q8.xlsm'
    out={'case':case,'status':'PREPARING','started_utc':utc(),'source_sha256':SOURCE_SHA,'fs_obtained':False,'target_met':False,'input_sha256':hashlib.sha256((ROOT/'inputs'/f'{case}.json').read_bytes()).hexdigest(),'scope_notes':rec['scope_notes']}
    try:
        assert not path.exists();shutil.copy2(TEMPLATE,path)
        app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
        out['excel_pid']=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        write(folder/'status.json',out)
        book=app.Workbooks.Open(str(path),UpdateLinks=0)
        call=lambda name,*args:app.Run("'"+book.Name+"'!"+name,*args)
        ws=book.Worksheets('要素定義');ws.Range('A7:T1006').ClearContents()
        for col,last,name in [('A','C','points'),('E','H','regions'),('J','T','faces')]:
            data=rec[name];rng=ws.Range(col+'7:'+last+str(6+len(data)))
            rng.Value2=tuple(tuple(v) for v in data)
            assert rng.Value2==tuple(tuple(v) for v in data),(case,name,rng.Value2)
        ws=book.Worksheets('材料データ');ws.Range('A5:Q1004').ClearContents()
        rng=ws.Range('A5:E'+str(4+len(rec['materials'])))
        rng.Value2=tuple(tuple(v) for v in rec['materials'])
        assert rng.Value2==tuple(tuple(v) for v in rec['materials']),(case,'materials',rng.Value2)
        for key,value in rec['settings'].items():setvalue(book,key,value)
        guide=book.Worksheets('P06N 実行案内');guide.Range('A1').Value2=case+' / G1R13 D03N/D07A修復の実Excel検証。入力表がこのケースの原本。'
        info=call('BatchDefinition').split('\t')
        assert [int(x) for x in info[:3]]==[len(rec['points']),len(rec['regions']),len(rec['faces'])],info
        assert abs(float(info[3])-rec['source_case']['area'])<=1e-10*rec['source_case']['area'],info
        assert info[4]==rec['settings']['ANALYSIS_POLICY'],info
        out['native_definition_check']=info
        out['actual_settings']={}
        ws=book.Worksheets('設定')
        for row in range(4,101):
            key=ws.Cells(row,4).Value2
            if key:out['actual_settings'][str(key)]=ws.Cells(row,2).Value2
        assert all(out['actual_settings'][k]==v for k,v in rec['settings'].items())
        call('BatchStagePath',str(folder/'stage.txt'))
        book.Save()
        out['status']='RUNNING';out['run_started_utc']=utc();write(folder/'status.json',out)
        print('START',case,'ExcelPID',out['excel_pid'],flush=True)
        tick=time.perf_counter();call('RPX_Run');out['solve_seconds']=time.perf_counter()-tick
        ws=book.Worksheets('解析結果');out['sheet_status']=ws.Range('A1').Value2
        bounds=ws.Range('B4:C4').Value2[0];out['lower_fs']=bounds[0];out['upper_fs']=bounds[1]
        out['relative_gap']=ws.Range('B9').Value2;out['element_counts']=ws.Range('B10:C10').Value2[0]
        out['summary_fields']=call('BatchResult',str(folder/'witness')).split('\t')
        out['error_number']=int(out['summary_fields'][0]);out['error']=out['summary_fields'][1]
        out['result_current']=out['summary_fields'][3]=='True';out['root_search_incomplete']=out['summary_fields'][4]=='True'
        out['root_width']=float(out['summary_fields'][5]);out['diagnostic_path']=out['summary_fields'][7]
        out['witness_q']=[int(v) for v in out['summary_fields'][8:10]];out['witness_phase']=out['summary_fields'][10:12]
        out['fs_values_present']=all(isinstance(v,(int,float)) and v>0 and v<1e20 for v in bounds) and bounds[0]<=bounds[1]
        out['status']='SOLVED_AWAITING_AUDIT' if out['fs_values_present'] and not out['error_number'] and out['result_current'] else 'FAILED'
        if out['status']=='SOLVED_AWAITING_AUDIT':
            from audit_batch import audit_case
            out['independent_audit']=audit_case(case,folder,rec,out)
            out['fs_obtained']=out['independent_audit']['pass']
            if not out['fs_obtained']:out['status']='AUDIT_FAILED'
            else:
                out=grade(out)
        book.Save();out['workbook_sha256']=hashlib.sha256(path.read_bytes()).hexdigest()
        out['workbook']=str(path.relative_to(ROOT))
    except Exception as e:
        out['status']='HARNESS_OR_EXPORT_ERROR';out['error']=repr(e);out['traceback']=traceback.format_exc()
    finally:
        out['seconds']=time.perf_counter()-started;out['finished_utc']=utc()
        write(folder/'status.json',out);write(ROOT/'results'/f'retry_{case}.json',out)
        if book is not None:
            try:book.Close(SaveChanges=False)
            except Exception:out['close_error']=traceback.format_exc()
        if app is not None:
            try:app.Quit()
            except Exception:out['quit_error']=traceback.format_exc()
        book=None;app=None;gc.collect();pythoncom.CoUninitialize()
    print('DONE',case,out['status'],'Fs',out.get('lower_fs'),out.get('upper_fs'),'seconds',round(out['seconds'],1),'error',out.get('error',''),flush=True)
    return out
def progress(ids):
    rows=[]
    for case in ids:
        path=ROOT/'cases'/case/'status.json'
        rows.append(json.loads(path.read_text(encoding='utf-8')) if path.exists() else {'case':case,'status':'NOT_RUN'})
    temp=ROOT/'results/progress.tmp'
    with temp.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(['case','status','fs_obtained','lower_fs','upper_fs','target_met','seconds','error'])
        for r in rows:w.writerow([r.get(k,'') for k in ['case','status','fs_obtained','lower_fs','upper_fs','target_met','seconds','error']])
    temp.replace(ROOT/'results/progress.csv')
    return rows
def batch(workers=4):
    ids=[p.stem for p in sorted((ROOT/'inputs').glob('*.json'))];assert len(ids)==25
    priority=['P06N','P06A','P08A','D01A','P09A','P10A','D02A','D03N']
    ordered=priority+[i for i in ids if i not in priority]
    with concurrent.futures.ProcessPoolExecutor(max_workers=workers,mp_context=multiprocessing.get_context('spawn')) as pool:
        pending={pool.submit(worker,case):case for case in ordered}
        while pending:
            done,_=concurrent.futures.wait(pending,timeout=10,return_when=concurrent.futures.FIRST_COMPLETED)
            for fut in done:
                case=pending.pop(fut)
                try:fut.result()
                except Exception:write(ROOT/'results'/f'{case}.json',{'case':case,'status':'WORKER_ERROR','error':traceback.format_exc()})
            progress(ids)
    rows=progress(ids)
    write(ROOT/'results/batch_summary.json',{'finished_utc':utc(),'all_cases_executed':all(r['status']!='NOT_RUN' for r in rows),'fs_obtained':sum(r.get('fs_obtained',False) for r in rows),'target_met':sum(r.get('target_met',False) for r in rows),'results':rows})
    assert hashlib.sha256(SOURCE.read_bytes()).hexdigest()==SOURCE_SHA
    print('ALL25_FINISHED',flush=True)
if __name__=='__main__':
    worker('D07A')
    worker('D03N')
