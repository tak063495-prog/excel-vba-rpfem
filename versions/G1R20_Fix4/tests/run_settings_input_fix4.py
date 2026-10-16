"""Owned Excel settings input contracts; no conic optimization.
TargetMet/event display probes use explicit synthetic sentinels, not certificates.
"""
from pathlib import Path
from fix4_artifact_binding import bind
import argparse, importlib.util, json, shutil, traceback
ROOT=Path(__file__).resolve().parents[1]; OLD=ROOT.parent/'G1R20_Fix3'
SPEC=importlib.util.spec_from_file_location('settings_session',OLD/'tests/run_saved_formula_excel.py')
U=importlib.util.module_from_spec(SPEC);SPEC.loader.exec_module(U)
TEMPLATE=OLD/'delivery/RPFEM_20261017_G1R20_Fix3.xlsm'
PROBE=ROOT/'tests/SettingsInputFix4Probe.bas'

def main():
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--negative',action='store_true');p.add_argument('--output',required=True,type=Path);a=p.parse_args()
    if not a.execute:raise SystemExit('Requires coordinated --execute')
    out=a.output.resolve();out.mkdir(parents=True,exist_ok=False)
    source=OLD if a.negative else ROOT;paths=sorted((source/'src/vba').glob('*.bas'));doc=source/'src/workbook/ThisWorkbook.code.txt'
    assert len(paths)==63
    monitored=[TEMPLATE,PROBE,Path(__file__),*paths,doc,*sorted((source/'import').glob('*.bas')),source/'import/ThisWorkbook.code.txt']
    monitored.append(ROOT/'tests/fix4_artifact_binding.py');before={str(x):U.sha(x) for x in monitored};snap=out/'source_snapshot';snap.mkdir()
    for item in [*paths,doc,PROBE]:shutil.copyfile(item,snap/item.name)
    shutil.copyfile(__file__,out/'runner_snapshot.py')
    rec=dict(status='PREPARING',negative=a.negative,sessions=[],cases=[],sha256_before=before,conic_optimizer_calls=0,fixture='Synthetic TargetMet/idle display probes; no numerical certificate generated')
    rec['final_artifact_binding']=bind(ROOT)
    def save():(out/'summary.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    tests=[]
    for value in [0,1,'OFF','ON',False,True,'FALSE','TRUE',2,'bad']:
        tests.append(('route_'+str(value),'SettingsRoute',[value],'route'))
    for value in [0,2,4,8,.4,1.6,7.6,3,True,'8']:
        tests.append(('q_'+str(value),'SettingsQuadrature',[value],'quadrature'))
    tests.extend([('duplicate','SettingsDuplicate',[],'duplicate'),('row110','SettingsOutOfRange',[],'row110'),('missing_dir','SettingsMissingSeismicKey',['missing_dir'],'missing_dir'),('missing_kh','SettingsMissingSeismicKey',['missing_kh'],'missing_kh')])
    if not a.negative:
        tests.extend([('row121','SettingsUnsupportedRow',[],'reject'),('row110_edit','SettingsExtendedEvent',[False],'event'),('row110_formula','SettingsExtendedEvent',[True],'event'),('gap_defaults','SettingsGapDefaults',[],'gap')])
        for scenario in ['both_missing','noop_formula','missing_dir_formula','missing_kh_formula']:tests.append(('seismic_'+scenario,'SettingsSeismicStable',[scenario],'seismic_stable'))
        for scenario in ['blank_formula_tail','error_key','numeric_key']:tests.append(('metadata_'+scenario,'SettingsMetadata',[scenario],'reject'))
        for scenario in ['notes_tail','canonical_keys']:tests.append(('metadata_'+scenario,'SettingsMetadata',[scenario],'accept'))
        tests.append(('mode_upper_canonical','SettingsGeneral',['MODE','FS'],'mode_canonical'))
        for key in ['MODE','GRAVITY']:tests.append(('error_'+key,'SettingsExcelError',[key],'reject'))
        invalid=[('MODE','UNKNOWN'),('GRAVITY','UNKNOWN'),('MODE',True),('GRAVITY',False),('QUADRATURE',1.5),('CYCLES',.5),('ADAPT_ACCURACY_ROUNDS',4.5),('ADAPT_ACCURACY_CAP',1536.4),('ADAPT_LOWER_FAN_ROUNDS',.4),('ADAPT_LOWER_TARGET',32.4),('ADAPT_REFINE_CAP',32.4),('HSD_LINEAR_CAPTURE_MAX',1.5),('HSD_LINEAR_TRACE',.5),('ADAPT_UPPER_FLOW_PASSES',.4),('RAW_LOWER_AUDIT',2),('HSD_SOLVE_SECONDS',0),('GAP',0),('GAP',2),('GAP',True),('SEISMIC_KH',True),('FS_TOL','bad'),('ADAPT_MESH_POLICY','unknown'),('DIRECT_DLL_SECONDS',601)]
        valid=[('GAP',.01),('GAP',.02),('GAP','0.02'),('HSD_STAGE_SECONDS',0),('HSD_STAGE_SECONDS',60),('SEISMIC_KH',None),('SEISMIC_KH','0.2'),('SEISMIC_DIR','左（-x）'),('LOWER_FIXED_RECOVERY',True),('ADAPT_ACCURACY_CAP',0),('ADAPT_LOWER_TARGET',0),('ADAPT_LOWER_FAN_ROUNDS',4),('ADAPT_ACCURACY_ROUNDS',20),('HSD_LINEAR_CAPTURE_MAX',0),('DIRECT_DLL_SECONDS',600),('QUADRATURE','8')]
        for key,value in invalid:tests.append(('invalid_'+key+'_'+str(value),'SettingsGeneral',[key,value],'reject'))
        for key,value in valid:tests.append(('valid_'+key+'_'+str(value),'SettingsGeneral',[key,value],'direction_canonical' if key=='SEISMIC_DIR' else 'accept'))
    save()
    try:
        for group in range(0,len(tests),20):
            path=out/('Settings_'+str(group//20)+'.xlsm');shutil.copyfile(TEMPLATE,path)
            with U.Session(rec,save,'settings_group_'+str(group//20)) as s:
                s.book=s.app.Workbooks.Open(str(path),UpdateLinks=0)
                for item in paths:
                    component=s.book.VBProject.VBComponents(item.stem)
                    if not a.negative:
                        component.CodeModule.DeleteLines(1,component.CodeModule.CountOfLines);component.CodeModule.AddFromString(U.code(item))
                    assert U.code_key(component.CodeModule.Lines(1,component.CodeModule.CountOfLines))==U.code_key(U.code(item)),item.name
                    component=None
                component=s.book.VBProject.VBComponents('ThisWorkbook')
                if not a.negative:
                    component.CodeModule.DeleteLines(1,component.CodeModule.CountOfLines);component.CodeModule.AddFromString(U.code(doc))
                assert U.code_key(component.CodeModule.Lines(1,component.CodeModule.CountOfLines))==U.code_key(U.code(doc));component=None
                component=s.book.VBProject.VBComponents.Add(1);component.Name=PROBE.stem;component.CodeModule.AddFromString(U.code(PROBE));component=None
                s.compile();s.app.EnableEvents=True
                for name,function,args,contract in tests[group:group+20]:
                    result=list(s.call(function,*args));case=dict(name=name,function=function,args=args,contract=contract,result=result);rec['cases'].append(case);save()
                    if a.negative:continue
                    if contract in ['reject','duplicate']:assert result[0]=='REJECTED' and result[1]==5,case
                    elif contract=='accept':assert result[0]=='ACCEPTED',case
                    elif contract=='route':
                        value=args[0]
                        if value in [2,'bad']:assert result[0]=='REJECTED',case
                        else:assert result[0]=='ACCEPTED' and result[1]==(str(value).upper() in ['1','ON','TRUE']),case
                    elif contract=='quadrature':
                        value=args[0]
                        if value in [.4,1.6,7.6,3,True]:assert result[0]=='REJECTED',case
                        else:assert result[0]=='ACCEPTED' and result[1]==(value==0),case
                    elif contract=='mode_canonical':assert result==['ACCEPTED','Fs'],case
                    elif contract=='direction_canonical':assert result==['ACCEPTED','左(-x)'],case
                    elif contract=='row110':assert result[0]=='ACCEPTED' and abs(result[1]-.2)<1e-12,case
                    elif contract=='missing_dir':assert result[0]=='ACCEPTED' and abs(result[1]-.2)<1e-12 and result[3:]==[1,1],case
                    elif contract=='missing_kh':assert result[0]=='ACCEPTED' and result[2]=='左(-x)' and result[3:]==[1,1],case
                    elif contract=='event':assert result[0]=='ACCEPTED' and result[1] is False and result[2] is None and abs(result[3]-.3)<1e-12,case
                    elif contract=='seismic_stable':
                        assert result[0]=='ACCEPTED' and result[4:6]==[1,1],case
                        if args[0]=='noop_formula':assert result[1]==-.2 and result[2]=='=0.2' and result[3]=='="左(-x)"' and result[6] is True,case
                        if args[0]=='missing_dir_formula':assert result[1]==.2 and result[2]=='=0.2',case
                        if args[0]=='missing_kh_formula':assert result[1]==0 and result[3]=='="左(-x)"',case
                    elif contract=='gap':assert result[0]=='ACCEPTED' and result[1]==.02 and result[2]=='2.0%' and result[3:]==[True,False,.01,.02,.02,.02],case
                    case['status']='PASS';save()
                s.book.Save()
        if a.negative:
            by={x['name']:x['result'] for x in rec['cases']}
            assert by['route_False'][1] is True and by['route_FALSE'][1] is True
            assert by['q_0.4'][1] is True and by['duplicate'][0]=='ACCEPTED' and by['row110'][1]==0
            assert by['missing_dir'][1]==0 and by['missing_dir'][3:]==[2,1]
            assert by['missing_kh'][2]=='右(+x)' and by['missing_kh'][3:]==[1,2]
        rec['status']='CONFIRMED_NEGATIVE' if a.negative else 'PASS'
    except BaseException as error:rec.update(status='FAIL',error=repr(error),traceback=traceback.format_exc());save();raise
    finally:
        rec['sha256_after']={str(x):U.sha(x) for x in monitored};rec['monitored_files_unchanged']=before==rec['sha256_after'];save()
    if not rec['monitored_files_unchanged'] or any(not x.get('owned_excel_exited') or x.get('dialogs') or x.get('monitor_errors') or x.get('hard_timeout') for x in rec['sessions']):
        rec.update(status='FAIL',post_run_verification_error='Ownership, immutability, dialogs, monitor or timeout failed');save()
        raise AssertionError(rec['post_run_verification_error'])
    print(json.dumps(dict(status=rec['status'],cases=len(rec['cases']),sessions=len(rec['sessions']),output=str(out)),ensure_ascii=False),flush=True)
if __name__=='__main__':main()
