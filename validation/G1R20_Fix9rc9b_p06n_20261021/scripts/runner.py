import os
for key in ['OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS']:os.environ[key]='1'
from pathlib import Path
import argparse,json,csv,shutil,hashlib,time,sys,gc,traceback,subprocess
from datetime import datetime,timezone
import win32com.client,win32process,pythoncom,threading,win32gui,win32con
from controls import start_controls
from prepare import ROOT,REPO,SOURCE,DLL,SOURCE_SHA,DLL_SHA,ORDER,sha,write,PRESET_MACRO,PRESET_EXPECTED,CASE_KEYS,same,TIMEOUTS,OVERALL_SECONDS,STOP_GRACE_SECONDS,PREPARING_LIMIT_SECONDS
from native_helpers import compile_book
from grade_results import grade
from audit_case import audit_case
TEMPLATE=ROOT/'prepared_template.xlsm'
def utc():return datetime.now(timezone.utc).isoformat()
def macros(path):
    from oletools.olevba import VBA_Parser
    parser=VBA_Parser(str(path))
    try:return {name:code.replace('\r\n','\n') for _,_,name,code in parser.extract_macros()}
    finally:parser.close()
def production_check(path):
    # Native Excel canonicalizes source when adding the evidence module.
    from verify_source import canonical
    before=macros(SOURCE);after=macros(path)
    assert set(after)==set(before)|{'BatchEvidence.bas'}
    for name in before:assert canonical(before[name])==canonical(after[name]),name
    return {'pass':True,'unchanged_original_modules':len(before),'only_added_module':'BatchEvidence','document_modules_unchanged':True}
def prepare_template():
    assert not TEMPLATE.exists() and sha(SOURCE)==SOURCE_SHA and sha(DLL)==DLL_SHA
    shutil.copyfile(SOURCE,TEMPLATE);app=None;book=None
    try:
        app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False
        book=app.Workbooks.Open(str(TEMPLATE),UpdateLinks=0)
        comp=book.VBProject.VBComponents.Add(1);comp.Name='BatchEvidence'
        text=(ROOT/'tests/BatchEvidence.bas').read_text(encoding='utf-8')
        comp.CodeModule.AddFromString('\r\n'.join(s for s in text.splitlines() if not s.startswith('Attribute VB_')))
        compile_book(app,book);book.Save()
        data={'compile':'PASS','source_sha256':SOURCE_SHA,'template_sha256':sha(TEMPLATE),'macro_security_unmodified':app.AutomationSecurity}
    finally:
        if book is not None:book.Close(False)
        if app is not None:app.Quit()
        book=None;app=None;gc.collect()
    data['integrity']=production_check(TEMPLATE);write(ROOT/'results/template_check.json',data)
    print('PREPARED TEMPLATE PASS',flush=True)
def settings(book):
    return {str(r[3]):r[1] for r in book.Worksheets('設定').Range('A4:D400').Value2 if r[3]}
SEISMIC_DEFAULTS={'SEISMIC_KH':0.0,'SEISMIC_DIR':'右(+x)'} # RPX_Seismic.bas RPX_Setting defaults when the row is absent
def configure(book,rec,call):
    ws=book.Worksheets('要素定義');ws.Range('A7:T1006').ClearContents()
    for first,last,name in [('A','C','points'),('E','H','regions'),('J','T','faces')]:
        data=tuple(tuple(row) for row in rec[name]);rng=ws.Range(first+'7:'+last+str(6+len(data)))
        rng.Value2=data;assert rng.Value2==data,(rec['case'],name)
    ws=book.Worksheets('材料データ');ws.Range('A5:Q1004').ClearContents()
    data=tuple(tuple(row) for row in rec['materials']);rng=ws.Range('A5:E'+str(4+len(data)))
    rng.Value2=data;assert rng.Value2==data
    # Product preset first (it also ensures missing setting rows); then only case-specific keys from the input.
    # Preset-owned keys from the G1R17 inputs (TARGET 1536, CYCLES 3, GAP 0.01) are never written.
    call(PRESET_MACRO)
    ws=book.Worksheets('設定');keys={str(row[3]):4+i for i,row in enumerate(ws.Range('A4:D400').Value2) if row[3]}
    absent={}
    for key in CASE_KEYS:
        if key not in keys and key in SEISMIC_DEFAULTS and rec['settings'][key]==SEISMIC_DEFAULTS[key]:
            absent[key]='row absent in delivery workbook; code default %r equals input'%(SEISMIC_DEFAULTS[key],);continue
        assert key in keys,('SETTING_MISSING',key)
        ws.Cells(keys[key],2).Value2=rec['settings'][key]
    # Fix8: optional extra settings (forced smoke variants only), written into free rows 4..120.
    for key,value in rec.get('extra_settings',{}).items():
        keys={str(row[3]):4+i for i,row in enumerate(ws.Range('A4:D120').Value2) if row[3]}
        if key not in keys:
            free=[4+i for i,row in enumerate(ws.Range('A4:D120').Value2) if not row[3] and not row[1]]
            assert free,('NO_FREE_SETTING_ROW',key);row=free[0];ws.Cells(row,4).Value2=key;keys[key]=row
        ws.Cells(keys[key],2).Value2=value
    actual=settings(book);assert all(actual[k]==rec['settings'][k] for k in CASE_KEYS if k not in absent)
    assert all(same(actual.get(k),v) for k,v in rec.get('extra_settings',{}).items()),('EXTRA_SETTINGS_NOT_APPLIED',rec.get('extra_settings'))
    actual['_absent_case_keys_using_default']=absent
    water=rec['water'];ws=book.Worksheets('浸透入力');assert ws.Range('Z1').Value2=='RPX_SEEP_LAYOUT_V2'
    ws.Range('A7:C206,E7:K206,M7:N206,P7:Q206').ClearContents()
    ws.Range('B3').Value2=ws.Range('Z6').Value2 # static groundwater level, mode=2
    ws.Range('B4').Value2=water['gamma_w'];ws.Range('D4').Value2=water['coupling'];ws.Range('F4').Value2=water['suction']
    ws.Range('C209').Value2=water['seep_target']
    if water['water_table']:ws.Range('M7:N8').Value2=tuple(tuple(row) for row in water['water_table'])
    for i,m in enumerate(rec['materials'],7):ws.Range(f'A{i}:C{i}').Value2=((m[0],1.,1.),)
    if water['coupling']:
        assert ws.Range('M7:N8').Value2==tuple(tuple(row) for row in water['water_table'])
    assert ws.Range('D4').Value2==water['coupling']
    # The delivery template contains levee seepage output. Dry cases must not
    # display or retain that unrelated field; wet cases regenerate their field.
    book.Worksheets('浸透結果').Cells.ClearContents()
    book.Worksheets('解析結果').Range('A1').Value2='未解析 / '+rec['case']
    for sheet in book.Worksheets:
        if '実行案内' in sheet.Name:sheet.Range('A1').Value2=rec['case']+' / G1R20 Fix9 rc8 P06 テスト（2%標準プリセット）。入力表がこのケースの原本。'
    return actual
def read_log(path):return list(csv.DictReader(path.open(encoding='utf-16'),delimiter='\t'))
def parts(value):return dict(p.split('=',1) for p in value.split(';') if '=' in p)
def log_summary(path):
    rows=read_log(path);ends=[r for r in rows if r['event']=='run_end']
    assert len(ends)==1
    audits=[r for r in rows if r['event']=='audit']
    rejected=[r for r in rows if 'audit' in r['event'] and ('status=FAIL' in r['detail'] or 'accepted=0' in r['detail'])]
    solves=[r for r in rows if r['event']=='solve_start']
    return {'analysis_seconds':float(ends[0]['elapsed_s']),'terminal':parts(ends[0]['detail']),
        'solve_count':len(solves),'audit_count':len(audits),'audit_rejections':rejected,
        'phase_seconds':{parts(r['detail'])['phase']:float(parts(r['detail'])['seconds']) for r in rows if r['event']=='run_phase'},
        'log_sha256':sha(path)}
def worker(case,smoke=False):
    if (ROOT/'HANDOFF').exists() and not smoke:
        # Batch hand-off to a new batch process: do nothing (no Excel, no status) and wait to be terminated.
        write(ROOT/'results'/f'handoff_worker_{case}.json',{'utc':utc(),'pid':os.getpid(),'case':case})
        while (ROOT/'HANDOFF').exists():time.sleep(2)
        raise SystemExit(75)
    pythoncom.CoInitialize();app=None;book=None;control_stop=None;tick=time.perf_counter()
    tag=('SMOKE_' if smoke else '')+case;folder=(ROOT/'harness_validation' if smoke else ROOT/'cases')/case
    folder.mkdir(parents=True,exist_ok=True)
    rec=json.loads((ROOT/'inputs'/f'{case}.json').read_text(encoding='utf-8'))
    path=ROOT/'workbooks'/f'G1R20F8_{tag}_Preset2pct.xlsm'
    out={'case':case,'tag':tag,'status':'PREPARING','started_utc':utc(),'pore_pressure':bool(rec['water']['coupling']),
        'source_sha256':SOURCE_SHA,'dll_sha256':DLL_SHA,'input_sha256':sha(ROOT/'inputs'/f'{case}.json'),
        'fs_obtained':False,'target_met':False,'smoke':smoke,'scope_notes':rec['scope_notes']}
    out['runner_sha256']=sha(Path(__file__))
    def persist():
        write(folder/'status.json',out)
        if not smoke:write(ROOT/'results'/f'{case}.json',out)
    try:
        assert not path.exists();assert sha(SOURCE)==SOURCE_SHA and sha(DLL)==DLL_SHA
        shutil.copyfile(TEMPLATE,path)
        app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False
        out['excel_pid']=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        out['excel_version']=app.Version;out['macro_security_unmodified']=app.AutomationSecurity;persist()
        control_stop=start_controls(app,out['excel_pid'],path.name,folder)
        book=app.Workbooks.Open(str(path),UpdateLinks=0)
        call=lambda name,*args:app.Run("'"+book.Name+"'!"+name,*args)
        absent=configure(book,rec,call)['_absent_case_keys_using_default'];actual=settings(book)
        mismatch={k:{'expected':v,'actual':actual.get(k)} for k,v in PRESET_EXPECTED.items() if k not in rec.get('extra_settings',{}) and not same(actual.get(k),v)}
        out['extra_settings']=rec.get('extra_settings',{})
        mismatch.update({k:{'expected':rec['settings'][k],'actual':actual.get(k)} for k in CASE_KEYS if k not in absent and actual.get(k)!=rec['settings'][k]})
        out['actual_settings']=actual;out['absent_case_keys_using_default']=absent;out['preset_macro']=PRESET_MACRO
        out['preset_check']={'pass':not mismatch,'checked_preset_keys':len(PRESET_EXPECTED),'checked_case_keys':len(CASE_KEYS),'mismatch':mismatch}
        persist();assert not mismatch,('PRESET_NOT_APPLIED',mismatch)
        info=str(call('BatchDefinition')).split('\t')
        assert [int(v) for v in info[:3]]==[len(rec['points']),len(rec['regions']),len(rec['faces'])],info
        assert math_close(float(info[3]),rec['source_case']['area']) and info[4]==rec['settings']['ANALYSIS_POLICY']
        out['native_definition']=info
        call('RPX_NativeSetMode','auto');out['native_start']=str(call('RPX_NativeInfo'))
        assert out['native_start'].startswith('DLL_READY'),out['native_start']
        call('RPX_SeepSetSilent',True);out['seep_seconds']=0.
        if rec['water']['coupling']:
            started=time.perf_counter();out['seep_message']=str(call('RPX_SeepRunSilent'));out['seep_seconds']=time.perf_counter()-started
            ws=book.Worksheets('浸透結果');assert ws.Range('K2').Value2 is True,'WATER_FIELD_NOT_READY'
            out['water_source_result']={'converged':ws.Range('K2').Value2,'nodes':ws.Range('K7').Value2,'elements':ws.Range('K6').Value2,'water_stamp':ws.Range('K9').Value2}
            assert '失敗' not in out['seep_message'] and 'ERR ' not in out['seep_message'],out['seep_message']
        call('RPX_SeepSetSilent',True)
        started=time.perf_counter();call('RPX_CreateMesh');out['mesh_seconds']=time.perf_counter()-started
        out['water_definition_check']=str(call('BatchWaterCheck',bool(rec['water']['coupling'])))
        assert out['water_definition_check'].startswith('PASS'),out['water_definition_check']
        call('BatchStagePath',str(folder/'stage.txt'))
        book.Save();out['status']='RUNNING';out['run_started_utc']=utc();persist()
        print('START',tag,'ExcelPID',out['excel_pid'],'TARGET',actual['TARGET'],'CYCLES',actual['CYCLES'],'GAP',actual['GAP'],'water',rec['water']['coupling'],flush=True)
        started=time.perf_counter();call('RPX_Run');out['com_analysis_seconds']=time.perf_counter()-started
        out['native_end']=str(call('RPX_NativeInfo'))
        assert out['native_end'].startswith('DLL_READY'),out['native_end']
        ws=book.Worksheets('解析結果');out['sheet_status']=ws.Range('A1').Value2
        out['lower_fs'],out['upper_fs']=ws.Range('B4:C4').Value2[0]
        out['relative_gap']=ws.Range('B9').Value2;out['element_counts']=ws.Range('B10:C10').Value2[0]
        out['summary_fields']=str(call('BatchResult',str(folder/'witness'))).split('\t')
        assert not out['summary_fields'][0].startswith('ERROR'),out['summary_fields']
        flags=out['summary_fields'];out['error_number']=int(flags[0]);out['error']=flags[1]
        out['result_current']=flags[3]=='True';out['root_search_incomplete']=flags[4]=='True';out['root_width']=float(flags[5])
        out['diagnostic_path']=flags[7];out['witness_q']=list(map(int,flags[8:10]));out['witness_phase']=flags[10:12]
        log=Path(out['diagnostic_path']);assert log.exists(),str(log)
        shutil.copyfile(log,folder/'analysis.tsv');out['log_summary']=log_summary(folder/'analysis.tsv')
        out['analysis_seconds']=out['log_summary']['analysis_seconds']
        out['fs_values_present']=all(isinstance(v,(int,float)) and 0<v<1e20 for v in [out['lower_fs'],out['upper_fs']]) and out['lower_fs']<=out['upper_fs']
        if out['fs_values_present'] and not out['error_number'] and out['result_current']:
            out['independent_audit']=audit_case(case,folder,rec,out)
            # Rejected intermediate candidates remain recorded. The final native
            # result and its independently audited witnesses determine Fs validity.
            out['fs_obtained']=out['independent_audit']['pass']
            out['status']='SOLVED_AWAITING_AUDIT' if out['fs_obtained'] else 'AUDIT_FAILED'
            if out['fs_obtained']:out=grade(out)
        else:out['status']='TIMEOUT_STOPPED' if (folder/'STOP_CASE').exists() else ('CANCELLED_BY_USER' if out['error_number']==18 and (ROOT/'STOP').exists() else 'FAILED')
        out['water_on_at_finish']=bool(call('BatchWaterOn'))
        assert out['water_on_at_finish']==bool(rec['water']['coupling'])
        out['production_source_check']=production_check_live(book)
        book.Save();out['workbook_sha256']=sha(path);out['workbook']=str(path.relative_to(ROOT))
    except Exception as e:
        out['status']='HARNESS_OR_EXPORT_ERROR';out['error']=repr(e);out['traceback']=traceback.format_exc()
        if book is not None:
            try:book.Save();out['workbook']=str(path.relative_to(ROOT));out['workbook_sha256']=sha(path)
            except Exception:out['save_error']=traceback.format_exc()
    finally:
        out['seconds_before_excel_close']=time.perf_counter()-tick;out['finished_utc']=utc();persist()
        if control_stop is not None:control_stop.set()
        if book is not None:
            try:book.Close(False)
            except Exception:out['close_error']=traceback.format_exc()
        if app is not None:
            try:app.Quit()
            except Exception:out['quit_error']=traceback.format_exc()
        book=None;app=None;gc.collect();pythoncom.CoUninitialize()
        out['total_worker_seconds']=time.perf_counter()-tick;persist()
    print('DONE',tag,out['status'],'Fs',out.get('lower_fs'),out.get('upper_fs'),'analysis_seconds',out.get('analysis_seconds'),'error',out.get('error',''),flush=True)
    return out
def math_close(a,b):return abs(a-b)<=1e-10*max(abs(b),1.)
def production_check_live(book):
    from verify_source import canonical
    original=macros(SOURCE)
    actual={c.Name:(c.CodeModule.Lines(1,c.CodeModule.CountOfLines) if c.CodeModule.CountOfLines else '') for c in book.VBProject.VBComponents}
    assert len(actual)==len(original)+1
    for filename,text in original.items():
        name=filename.rsplit('.',1)[0];assert canonical(text)==canonical(actual[name]),name
    return {'pass':True,'original_modules_unchanged':len(original),'only_added_module':'BatchEvidence'}
COLUMNS=['case','status','analysis_seconds','timeout_seconds','lower_fs','upper_fs','relative_gap','fs_obtained','target_met','error']
def _read_json_retry(path,tries=20):
    # rev3: tolerate transient Windows sharing violations / partial reads while the worker replaces status.json
    for i in range(tries):
        try:return json.loads(path.read_text(encoding='utf-8'))
        except (PermissionError,json.JSONDecodeError):
            if i==tries-1:raise
            time.sleep(0.25)
def progress():
    rows=[]
    for case in ORDER:
        path=ROOT/'cases'/case/'status.json'
        rows.append(_read_json_retry(path) if path.exists() else {'case':case,'status':'NOT_RUN'})
    for r in rows:r.setdefault('timeout_seconds',TIMEOUTS[r['case']]['timeout_seconds'])
    path=ROOT/'results/progress.csv';temp=path.with_suffix('.tmp')
    with temp.open('w',encoding='utf-8-sig',newline='') as f:
        w=csv.writer(f);w.writerow(COLUMNS)
        for row in rows:w.writerow([row.get(k,'') for k in COLUMNS])
    temp.replace(path)
    write(ROOT/'results/batch_summary.json',{'updated_utc':utc(),'count':len(ORDER),'all_cases_executed':all(r['status'] not in ['NOT_RUN','PREPARING','RUNNING'] for r in rows),
        'fs_obtained':sum(bool(r.get('fs_obtained',False)) for r in rows),'target_met':sum(bool(r.get('target_met',False)) for r in rows),
        'timeouts':[r['case'] for r in rows if str(r['status']).startswith('TIMEOUT')],'results':rows})
    return rows
def excel_pids():
    out=subprocess.run(['tasklist','/FI','IMAGENAME eq EXCEL.EXE','/FO','CSV','/NH'],capture_output=True,text=True,errors='replace').stdout
    return {int(l.split('","')[1]) for l in out.splitlines() if l.startswith('"EXCEL')}
def mark(case,updates):
    folder=ROOT/'cases'/case;path=folder/'status.json'
    data=json.loads(path.read_text(encoding='utf-8')) if path.exists() else {'case':case}
    data.update(updates);write(path,data);write(ROOT/'results'/f'{case}.json',data);return data
def run_case(case,foreign,deadline):
    folder=ROOT/'cases'/case;limit=TIMEOUTS[case]['timeout_seconds']
    stopfile=folder/'STOP_CASE'
    log=(ROOT/'results'/f'batch_stdout_{case}.txt').open('a',encoding='utf-8')
    proc=subprocess.Popen([sys.executable,'-X','utf8',str(Path(__file__).resolve()),'case',case],cwd=REPO,stdout=log,stderr=subprocess.STDOUT)
    spawned=time.time();run_t0=None;stop_t=None;killed=None;effective=limit
    while proc.poll() is None:
        time.sleep(5);progress();now=time.time()
        st={}
        try:st=json.loads((folder/'status.json').read_text(encoding='utf-8'))
        except Exception:pass
        if run_t0 is None and st.get('status')=='RUNNING':
            run_t0=now;effective=min(limit,max(60,deadline-now))
        over_prep=run_t0 is None and now-spawned>PREPARING_LIMIT_SECONDS
        over_run=run_t0 is not None and now-run_t0>effective
        if (over_prep or over_run or (ROOT/'STOP').exists()) and stop_t is None:
            stop_t=now
            if not (ROOT/'STOP').exists():write(stopfile,{'utc':utc(),'reason':'PREPARING_LIMIT' if over_prep else 'TIMEOUT','limit_seconds':effective})
        if stop_t is not None and now-stop_t>(60 if over_prep else STOP_GRACE_SECONDS) and killed is None:
            pid=st.get('excel_pid');killed={'utc':utc(),'excel_pid':pid,'excel_killed':False}
            if pid and pid in excel_pids() and pid not in foreign:
                r=subprocess.run(['taskkill','/PID',str(pid),'/F'],capture_output=True,text=True,errors='replace')
                killed['excel_killed']=r.returncode==0;killed['taskkill']=(r.stdout+r.stderr).strip()
            try:proc.wait(90)
            except subprocess.TimeoutExpired:proc.kill();proc.wait();killed['worker_killed']=True
    log.close()
    timing={'timeout_seconds':limit,'effective_timeout_seconds':effective,'best_prior_seconds':TIMEOUTS[case]['best_prior_seconds'],
        'other_excel_pids_at_start':sorted(foreign),'runner_wall_seconds':time.time()-spawned,'run_wall_seconds':(time.time()-run_t0) if run_t0 else None,'worker_exit_code':proc.returncode}
    st=json.loads((folder/'status.json').read_text(encoding='utf-8')) if (folder/'status.json').exists() else {}
    if killed is not None:
        mark(case,{**timing,'worker_status_before_kill':st.get('status'),'worker_error_before_kill':st.get('error'),
            'status':'TIMEOUT_KILLED','forced_kill':killed,'fs_obtained':False,'target_met':False,'finished_utc':utc(),
            'error':'RPX_Stop did not finish within %d s; runner-launched Excel terminated'%STOP_GRACE_SECONDS})
    elif st.get('status') in (None,'PREPARING','RUNNING'):
        mark(case,{**timing,'worker_status_before_exit':st.get('status'),'status':'WORKER_DIED','fs_obtained':False,'target_met':False,'finished_utc':utc()})
    else:
        extra={'timed_out':stopfile.exists()}
        if stopfile.exists() and st.get('status')!='TIMEOUT_STOPPED':
            extra.update({'worker_status_before_timeout_mark':st.get('status'),'status':'TIMEOUT_STOPPED','fs_obtained':False,'target_met':False})
        mark(case,{**timing,**extra})
    progress()
def batch():
    assert TEMPLATE.exists()
    startfile=ROOT/'results/batch_start.json'
    if startfile.exists():start=json.loads(startfile.read_text(encoding='utf-8'))
    else:
        start={'utc':utc(),'epoch':time.time(),'overall_limit_seconds':OVERALL_SECONDS};write(startfile,start)
    deadline=start['epoch']+OVERALL_SECONDS
    foreign=excel_pids()
    write(ROOT/'results'/f'other_excel_at_batch_start_{int(time.time())}.json',{'utc':utc(),'excel_pids_not_launched_by_runner':sorted(foreign),
        'note':'Never terminated by the runner; may share CPU with the benchmark'})
    for case in ORDER:
        if (ROOT/'STOP').exists():
            write(ROOT/'results/user_interruption.json',{'utc':utc(),'source':'STOP marker; cooperative RPX_Stop'});progress();return
        existing=ROOT/'cases'/case/'status.json'
        if existing.exists():
            data=json.loads(existing.read_text(encoding='utf-8'))
            rerun_short=str(data['status']).startswith('TIMEOUT') and (data.get('effective_timeout_seconds') or 0)<TIMEOUTS[case]['timeout_seconds']-1
            if rerun_short:
                stamp=str(int(time.time()));(ROOT/'cases'/case).rename(ROOT/'cases'/(case+'_short_timeout_'+stamp))
                wb=ROOT/'workbooks'/f'G1R20F8_{case}_Preset2pct.xlsm'
                if wb.exists():wb.rename(wb.with_name(wb.stem+'_short_timeout_'+stamp+'.xlsm'))
                rj=ROOT/'results'/f'{case}.json'
                if rj.exists():rj.rename(rj.with_name(case+'_short_timeout_'+stamp+'.json'))
            elif data['status'] in ['RUNNING','PREPARING']:
                pid=data.get('excel_pid')
                if pid and pid in excel_pids():raise SystemExit('LIVE_UNFINISHED_CASE '+case)
                old=ROOT/'cases'/(case+'_stale_'+str(int(time.time())));(ROOT/'cases'/case).rename(old)
                wb=ROOT/'workbooks'/f'G1R20F8_{case}_Preset2pct.xlsm'
                if wb.exists():wb.rename(wb.with_name(wb.stem+'_stale_'+str(int(time.time()))+'.xlsm'))
            else:continue
        if time.time()>deadline:
            mark(case,{'status':'NOT_RUN_OVERALL_LIMIT','fs_obtained':False,'target_met':False});continue
        run_case(case,foreign|set(),deadline)
        if case=='D08A':
            d=json.loads((ROOT/'cases/D08A/status.json').read_text(encoding='utf-8'))
            ok=d.get('fs_obtained') and d.get('preset_check',{}).get('pass') and d['status'] in ['PASS_TARGET','FS_AVAILABLE_PARTIAL']
            write(ROOT/'results/smoke_D08A.json',{'utc':utc(),'pass':bool(ok),'status':d.get('status'),'lower_fs':d.get('lower_fs'),'upper_fs':d.get('upper_fs'),
                'relative_gap':d.get('relative_gap'),'analysis_seconds':d.get('analysis_seconds'),'preset_check':d.get('preset_check')})
            if not ok:
                write(ROOT/'results/HALTED.json',{'utc':utc(),'reason':'D08A smoke gate failed','status':d.get('status'),'error':d.get('error')})
                progress();print('HALTED_SMOKE_FAILED',flush=True);return
        ours=json.loads((ROOT/'cases'/case/'status.json').read_text(encoding='utf-8')).get('excel_pid')
        foreign=foreign|(excel_pids()-{ours})
    rows=progress()
    write(ROOT/'results/ALL_FINISHED.json',{'utc':utc(),'elapsed_hours':(time.time()-start['epoch'])/3600,
        'statuses':{r['case']:r['status'] for r in rows},'source_sha_unchanged':sha(SOURCE)==SOURCE_SHA,'dll_sha_unchanged':sha(DLL)==DLL_SHA})
    print('ALL_FINISHED',flush=True)
if __name__=='__main__':
    if sys.argv[1]=='prepare':prepare_template()
    elif sys.argv[1]=='smoke':
        result=worker('D08A',True);assert result['fs_obtained'] and result['status'] in ['PASS_TARGET','FS_AVAILABLE_PARTIAL'],result
    elif sys.argv[1]=='case':worker(sys.argv[2])
    elif sys.argv[1]=='batch':batch()
