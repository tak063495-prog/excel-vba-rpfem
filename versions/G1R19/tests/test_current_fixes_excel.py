"""Serial owned Excel cold runs using the current full source and latest fixes.

Workbooks, snapshots and logs are new outputs. No active/user Excel is attached.
The test copy adds read-only state checkpoint and native-worker telemetry calls.
"""
from pathlib import Path
import argparse
import codecs
import collections
import gc
import hashlib
import json
import os
import shutil
import struct
import sys
import threading
import time
import uuid

for key in ['OMP_NUM_THREADS','OPENBLAS_NUM_THREADS','MKL_NUM_THREADS','RAYON_NUM_THREADS']:
    os.environ[key]='1'
import psutil
import win32com.client
import win32con
import win32gui
import win32process

ROOT=Path(__file__).resolve().parents[1]
REPO=ROOT.parents[1]
sys.path.insert(0,str(ROOT/'tools'))
from current_sources import module_paths, sha


def source_text(path):
    return '\r\n'.join(line for line in path.read_text(encoding='utf8').splitlines() if not line.startswith('Attribute VB_'))


def atomic_json(path,value):
    tmp=path.with_suffix(path.suffix+'.tmp')
    tmp.write_text(json.dumps(value,indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8')
    tmp.replace(path)


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--source-root',type=Path,required=True)
    parser.add_argument('--case',choices=['P06A','P06N'],default='P06A')
    parser.add_argument('--preflight-only',action='store_true',default=True)
    parser.add_argument('--precision-from-start',action='store_true')
    parser.add_argument('--accuracy-cap',type=int,default=24000)
    parser.add_argument('--lower-target',type=int,default=4000)
    args=parser.parse_args()
    if not 1000<=args.accuracy_cap<=50000 or not (args.lower_target==0 or 16<=args.lower_target<=args.accuracy_cap):
        raise ValueError('Explicit accuracy budgets must fit the existing 50000-element safety cap')
    sources=module_paths(args.source_root)
    if args.precision_from_start and 'Private Function PrecisionFromStart()' not in sources['RPX_Adapt'].read_text(encoding='utf8'):
        raise ValueError('Requested q8 start needs its explicitly generated source option')
    prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
    if prior:raise RuntimeError(f'Another Excel is running; serial ownership required: {prior}')
    token=uuid.uuid4().hex
    folder=ROOT/'results'/('current_'+('preflight' if args.preflight_only else args.case)+'_'+str(time.time_ns()))
    folder.mkdir(parents=True);(folder/'checkpoints').mkdir()
    workbook=folder/'CurrentProduction.xlsm'
    shutil.copyfile(ROOT.parent/'G1R18/delivery/RPFEM_20261014_G1R18.xlsm',workbook)
    for path in [ROOT/'delivery/rpfem_native_x64.dll']:
        shutil.copyfile(path,folder/path.name)
    record={'status':'PREPARING','run_token':token,'case':args.case,'kind':'preflight' if args.preflight_only else 'cold_run',
            'source_root':str(args.source_root.resolve()),'source_sha256':{n:sha(p) for n,p in sources.items()},
            'dll_sha256':{p.name:sha(p) for p in folder.glob('*.dll')},'harness_sha256':sha(__file__),
            'observer_sha256':sha(Path(__file__).with_name('OwnedProductionProbe.bas')),
            'solver_threads':1,'excel_concurrency':1,'prior_excel_pids':prior,'output':str(folder),
            'published_artifacts_modified':False,'imported_field_or_root':False,'dialogs':[]}
    shutil.copyfile(__file__,folder/'harness_snapshot.py')
    shutil.copyfile(Path(__file__).with_name('OwnedProductionProbe.bas'),folder/'observer_snapshot.bas')
    save=lambda:atomic_json(folder/'summary.json',record)
    save()
    app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    assert pid not in prior
    proc=psutil.Process(pid)
    record.update(pid=pid,process_creation_time=proc.create_time(),harness_pid=os.getpid());save()
    stop=threading.Event()
    def watch():
        cursors={};events=collections.Counter();last=[];peak=0
        while not stop.wait(2):
            try:
                memory=proc.memory_info();peak=max(peak,memory.rss)
                for path in list((folder/'RPFEM_logs').glob('*.tsv'))+list(folder.glob('*.tsv')):
                    if path.name.endswith('.meta.tsv'):continue
                    state=cursors.get(path)
                    if state is None:
                        with path.open('rb') as source:head=source.read(2)
                        if len(head)<2:continue
                        enc='utf-16' if head in [b'\xff\xfe',b'\xfe\xff'] else 'utf-8-sig'
                        state=[0,codecs.getincrementaldecoder(enc)('replace'),''];cursors[path]=state
                    with path.open('rb') as source:source.seek(state[0]);data=source.read();state[0]=source.tell()
                    text=state[2]+state[1].decode(data);lines=text.split('\n');state[2]=lines.pop()
                    for line in lines:
                        parts=line.rstrip('\r').split('\t')
                        markers=[part for part in parts if part.startswith(('accuracy_dll_','owned_accuracy_','lower_cache_','lower_prepared_','candidate_end','accuracy_round','own_quadrature_','fs_root')) and len(part)<100]
                        if markers:
                            event=markers[0];events[event]+=1
                            if 'accepted=True' in line:events[event+':accepted']+=1
                            if 'accepted=False' in line:events[event+':rejected']+=1
                            if 'original_evidence' in line:events['cache_original_evidence_drop']+=1
                            last.append(line[-3500:]);last=last[-12:]
                heart={'run_token':token,'pid':pid,'process_creation_time':record['process_creation_time'],
                       'time_unix':time.time(),'rss_bytes':memory.rss,'peak_observed_rss_bytes':peak,
                       'cpu_seconds':sum(proc.cpu_times()[:2]),'available_ram_bytes':psutil.virtual_memory().available,
                       'ready_checkpoints':len(list((folder/'checkpoints').glob('*.meta.tsv'))),
                       'events':dict(events),'last_events':last,'resource_advisory':peak>12*1024**3}
                atomic_json(folder/'heartbeat.json',heart)
                def visit(hwnd,_):
                    if win32process.GetWindowThreadProcessId(hwnd)[1]!=pid or not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd)!='#32770':return
                    children=[]
                    win32gui.EnumChildWindows(hwnd,lambda child,_:children.append((child,win32gui.GetWindowText(child))),None)
                    texts=[text for _,text in children]
                    if texts not in record['dialogs']:record['dialogs'].append(texts)
                    for child,text in children:
                        if text in ['OK','&OK'] or (win32gui.GetWindowText(hwnd)=='Microsoft Visual Basic' and ('終了' in text or '&End' in text)):
                            win32gui.PostMessage(child,win32con.BM_CLICK,0,0)
                win32gui.EnumWindows(visit,None)
            except (psutil.NoSuchProcess,psutil.AccessDenied):break
            except Exception as error:
                atomic_json(folder/'monitor_error.json',{'error':repr(error),'time_unix':time.time()})
    watcher=threading.Thread(target=watch,daemon=True);watcher.start()
    book=component=module=ws=control=None
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
        app.MultiThreadedCalculation.Enabled=False
        book=app.Workbooks.Open(str(workbook),UpdateLinks=0)
        for name,path in sources.items():
            component=book.VBProject.VBComponents(name);module=component.CodeModule
            module.DeleteLines(1,module.CountOfLines);code=source_text(path)
            if name=='RPX_Adapt':
                old='    state = pending: SaveState = True'
                assert code.count(old)==1
                code=code.replace(old,old+'\r\n    OwnedProductionCheckpoint state, side',1)
                if args.preflight_only:
                    old='    If mode = "Fs" Then\r\n        RFsActive = True'
                    assert code.count(old)==1
                    code=code.replace(old,'    OwnedObserveQuality RSubdivisions, RFsTolerance\r\n'+old,1)
                    code+='''
Public Function OwnedQualityPreflight() As Variant
    Dim previous As Boolean, number As Long, message As String
    OwnedObservedQ = 0: OwnedObservedTolerance = 0
    previous = adaptCoarseSearch: adaptCoarseSearch = True: OwnedQualityArmed = True
    On Error GoTo Captured
    Evaluate "Fs", "quality_probe"
Captured:
    number = Err.Number: message = Err.Description
    OwnedQualityArmed = False: adaptCoarseSearch = previous
    If number = 18 And OwnedObservedQ > 0 And message = "ANALYSIS_CANCELLED" Then
        OwnedQualityPreflight = Array(OwnedObservedQ, OwnedObservedTolerance)
    Else
        OwnedQualityPreflight = Array("REJECT", number, message)
    End If
End Function
'''
            if name=='RPX_AccuracyNative':
                old='    rc = NStart(handle): If rc <> 0 Then GoTo Cleanup'
                assert code.count(old)==1
                code=code.replace(old,old+'\r\n    RPX_DiagEvent "owned_accuracy_worker_started", "mode=" & pr.mode & ";Fs=" & pr.factor & ";n=" & pr.n & ";m=" & pr.m & ";threads=1"',1)
                old='    If Not valid Then Erase answer: Erase modifiedDual'
                assert code.count(old)==1
                code=code.replace(old,old+'\r\n    RPX_DiagEvent "owned_accuracy_worker_end", "mode=" & pr.mode & ";finite_candidate=" & valid & ";solver_status=" & AccuracyNativeSolverStatus & ";iterations=" & AccuracyNativeIterations',1)
            module.AddFromString(code)
        extra={'OwnedProductionProbe':Path(__file__).with_name('OwnedProductionProbe.bas'),
               'AccuracyExport':REPO/'validation/G1R12_native25_20261009/scripts/tests/BatchEvidence.bas'}
        if args.preflight_only:extra['Issue15Probe']=ROOT.parent/'G1R18_Issue15/tests/Issue15Probe.bas'
        for name,path in extra.items():
            component=book.VBProject.VBComponents.Add(1);component.Name=name
            component.CodeModule.AddFromString(source_text(path).replace('Private Sub BatchState(','Public Sub BatchState('))
        book.Activate();book.VBProject.VBComponents('OwnedProductionProbe').CodeModule.CodePane.Show()
        control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled:control.Execute();time.sleep(.3)
        if control.Enabled:raise RuntimeError('Current full VBA compile failed')
        app.VBE.MainWindow.Visible=False;record['compile']='PASS'
        record['loaded_code_sha256']={c.Name:hashlib.sha256(c.CodeModule.Lines(1,c.CodeModule.CountOfLines).encode('utf8')).hexdigest() for c in book.VBProject.VBComponents if c.Type==1}
        call=lambda name,*values:app.Run("'"+book.Name+"'!"+name,*values)
        save();print(str(folder)+' compile PASS ownedPID='+str(pid),flush=True)
        if args.preflight_only:
            result=call('OwnedCachePreflight',str(folder/'cache.tsv'));record['cache_preflight']=result;save()
            assert str(result).startswith('PASS'),result
            if 'Private Function PrecisionFromStart()' in sources['RPX_Adapt'].read_text(encoding='utf8'):
                call('OwnedSetting','QUADRATURE',8);call('OwnedSetting','FS_TOL',.0001)
                record['actual_coarse_quality']={}
                for flag,expected in [(0,[4,.001]),(1,[8,.0001])]:
                    call('OwnedSetting','ADAPT_PRECISION_FROM_START',flag)
                    observed=list(call('OwnedQualityPreflight'))
                    assert observed==expected,(flag,observed)
                    record['actual_coarse_quality'][str(flag)]=observed
                call('OwnedSetting','ADAPT_PRECISION_FROM_START',2)
                observed=list(call('OwnedQualityPreflight'))
                assert observed[0]=='REJECT' and 'INVALID_ADAPT_PRECISION_FROM_START' in observed[-1],observed
                record['actual_coarse_quality']['invalid_setting']=observed
                call('OwnedSetting','ADAPT_PRECISION_FROM_START',0)
                save()
            record['seep_preflight']={}
            for kind in ['validation','unconverged','success']:
                for warm in [False,True]:
                    result=call('Issue15Case',kind,warm);record['seep_preflight'][kind+'/'+str(warm)]=result
                    assert (';after=REJECTED:' in result) if kind!='success' else (';after=ACCEPTED' in result),result
                    save()
            record['status']='PASS';book.Save()
        else:
            input_path=REPO/'validation/accuracy_20261013/python/data'/(args.case+'.json')
            inputs=json.loads(input_path.read_text(encoding='utf8'))
            if inputs.get('water',{}).get('coupling',0)!=0:raise ValueError('Cold P06 baseline expects original dry case')
            record['input_sha256']=sha(input_path);record['input_path']=str(input_path)
            shutil.copyfile(input_path,folder/'input_record.json')
            call('RPX_Setup')
            for sheet,ranges in [('要素定義',[('A','C','points',7),('E','H','regions',7),('J','T','faces',7)]),('材料データ',[('A','E','materials',5)])]:
                ws=book.Worksheets(sheet);ws.Range('A7:T1006' if sheet=='要素定義' else 'A5:Q1004').ClearContents()
                for col,last,key,first in ranges:
                    rows=tuple(tuple(row) for row in inputs[key]);rng=ws.Range(f'{col}{first}:{last}{first+len(rows)-1}')
                    rng.Value2=rows;assert rng.Value2==rows
                rng=None
            for key,value in inputs['settings'].items():call('OwnedSetting',key,value)
            before={name:book.Worksheets(name).UsedRange.Value2 for name in ['要素定義','材料データ']}
            call('RPX_AccuracyPreset')
            assert before=={name:book.Worksheets(name).UsedRange.Value2 for name in before}
            for key,value in {'ADAPT_MESH_POLICY':'ROBUST_REFINE','LOWER_MODEL_CACHE':1,'TEST_INJECTION':0,'HSD_CAPTURE':0,'DIAGNOSTIC_MODE':1,'ACCURACY_DLL_CAPTURE':1}.items():call('OwnedSetting',key,value)
            call('OwnedSetting','ADAPT_PRECISION_FROM_START',int(args.precision_from_start))
            call('OwnedSetting','ADAPT_ACCURACY_CAP',args.accuracy_cap)
            call('OwnedSetting','ADAPT_LOWER_TARGET',args.lower_target)
            book.Worksheets('浸透入力').Range('D4').Value2=0
            settings={str(row[2]):row[0] for row in book.Worksheets('設定').Range('B4:D999').Value2 if row[2]}
            assert settings['ACCURACY_DLL_RECOVERY']==1 and settings['ADAPT_CERT_QUADRATURE']==8 and settings['QUADRATURE']==8
            assert settings['ANALYSIS_POLICY']==inputs['settings']['ANALYSIS_POLICY'] and settings['LOWER_MODEL_CACHE']==1
            assert settings['ADAPT_ACCURACY_CAP']==args.accuracy_cap and settings['ADAPT_LOWER_TARGET']==args.lower_target
            definition=str(call('AccuracyExport.BatchDefinition')).split('\t')
            assert list(map(int,definition[:3]))==[len(inputs[key]) for key in ['points','regions','faces']]
            assert definition[-1]==inputs['settings']['ANALYSIS_POLICY']
            assert abs(float(definition[3])-inputs['source_case']['area'])<1e-8
            record.update(settings=settings,native_definition=definition,status='RUNNING',started_unix=time.time())
            book.Save();record['configured_workbook_sha256']=sha(workbook);save()
            result=call('OwnedProductionRun',str(folder/'production.tsv'),str(folder/'checkpoints'))
            record['result_variant']=list(result);record['finished_unix']=time.time()
            record['result_double_hex']={str(i):struct.pack('<d',v).hex() for i,v in enumerate(result) if isinstance(v,float)}
            record['precision_target_met']=result[0]=='AUDITED' and bool(result[1])
            record['status']='AUDITED' if result[0]=='AUDITED' else 'PARTIAL'
            if result[0]=='FAIL':record['status']='FAIL'
            book.Save();save()
        print(json.dumps({k:record[k] for k in ['status','pid','run_token','output']},ensure_ascii=False),flush=True)
    except BaseException as error:
        record['status']='FAIL';record['error']=repr(error);save();raise
    finally:
        try:
            if book is not None:book.Close(SaveChanges=False)
            app.Quit()
        finally:
            stop.set();watcher.join(timeout=3)
            book=component=module=ws=control=app=None;gc.collect()
        for _ in range(50):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        record['owned_excel_exited']=not psutil.pid_exists(pid)
        record['files']={p.relative_to(folder).as_posix():{'bytes':p.stat().st_size,'sha256':sha(p)} for p in folder.rglob('*') if p.is_file() and p.name not in ['summary.json','summary.json.tmp']}
        save()


if __name__=='__main__':main()
