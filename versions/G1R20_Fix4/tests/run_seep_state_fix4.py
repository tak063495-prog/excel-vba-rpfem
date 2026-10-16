"""Owned Excel: failed seep retires dependent soil display and restores application state.
Synthetic main display/flags are explicit; real public mode2 seep producer runs;
no conic optimization or numerical soil certificate is claimed by this test.
"""
from pathlib import Path
from fix4_artifact_binding import bind
import argparse, importlib.util, json, shutil, traceback
ROOT=Path(__file__).resolve().parents[1]
OLD=ROOT.parent/'G1R20_Fix3'
SPEC=importlib.util.spec_from_file_location('saved_formula_util',OLD/'tests/run_saved_formula_excel.py')
U=importlib.util.module_from_spec(SPEC);SPEC.loader.exec_module(U)
TEMPLATE=OLD/'delivery/RPFEM_20261017_G1R20_Fix3.xlsm'
PROBE=ROOT/'tests/WaterSourceFix3Probe.bas'
HELPER='Option Explicit\nPublic Sub SeedSyntheticMain(ByVal wet As Boolean)\n    RPX_InternalWrite = True\n    With ThisWorkbook.Worksheets("解析結果")\n        .Range("A1").Value2 = "TEST_SYNTHETIC_CERTIFIED_DISPLAY"\n        .Range("B4").Value2 = 1.11\n        .Range("C4").Value2 = 2.22\n    End With\n    RPX_ResultCurrent = True: RPX_MeshCurrent = True: RPX_WaterApplied = wet\n    RPX_InternalWrite = False\n    Call RPX_InputObserve\n    Call RPX_InputCommit\nEnd Sub\nPublic Function ObserveSoilFlags() As Variant\n    ObserveSoilFlags = Array(RPX_ResultCurrent, RPX_MeshCurrent, RPX_Busy, RPX_InternalWrite, XCancel)\nEnd Function\nPublic Function RunWithIncomingState(ByVal screen As Boolean, ByVal cancelKey As Long, ByVal suspend As Boolean) As Variant\n    Dim message As String\n    Application.ScreenUpdating = screen: Application.EnableCancelKey = cancelKey: RPX_WaterSuspend = suspend\n    message = RPX_SeepRunSilent()\n    RunWithIncomingState = Array(message, Application.ScreenUpdating, Application.EnableCancelKey, RPX_WaterSuspend)\nEnd Function\n'


def main():
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--negative',action='store_true');p.add_argument('--output',required=True,type=Path);a=p.parse_args()
    if not a.execute:raise SystemExit('Requires coordinated --execute')
    out=a.output.resolve();out.mkdir(parents=True,exist_ok=False)
    source_root=OLD if a.negative else ROOT
    sources=sorted((source_root/'src/vba').glob('*.bas'));doc=source_root/'src/workbook/ThisWorkbook.code.txt'
    assert len(sources)==63
    monitored=[TEMPLATE,PROBE,Path(__file__),*sources,doc,*sorted((source_root/'import').glob('*.bas')),source_root/'import/ThisWorkbook.code.txt']
    monitored.append(ROOT/'tests/fix4_artifact_binding.py');before={str(x):U.sha(x) for x in monitored};snapshot=out/'source_snapshot';snapshot.mkdir()
    for source in [*sources,doc,PROBE]:shutil.copyfile(source,snapshot/source.name)
    shutil.copyfile(__file__,out/'runner_snapshot.py');(out/'StateProbe.bas').write_text(HELPER,encoding='utf8')
    rec=dict(status='PREPARING',negative=a.negative,sessions=[],cases=[],fixture='Real public mode2 seep producer, synthetic main display/flags/source proof; conic optimizer0, no numerical soil certificate',sha256_before=before)
    rec['final_artifact_binding']=bind(ROOT)
    def save():(out/'summary.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    def observe(s):
        result=s.book.Worksheets('解析結果');water=s.book.Worksheets('浸透結果');inp=s.book.Worksheets('浸透入力')
        val=dict(flags=list(s.call('ObserveSoilFlags')),status=result.Range('A1').Value2,pair=list(result.Range('B4:C4').Value2[0]),water_k2=water.Range('K2').Value2,water_k9=water.Range('K9').Value2,coupling=inp.Range('D4').Value2,coupling_formula=inp.Range('D4').Formula,journal=s.book.CustomDocumentProperties('RPX_SeepResultWriteState').Value)
        result=water=inp=None
        return val
    save()
    cases=['protected_output_coupling_formula'] if a.negative else ['protected_output_coupled','protected_output_coupling_formula','protected_output_dry','success_coupled','secondary_protected_soil_primary_gamma']
    try:
        for index,label in enumerate(cases):
            path=out/(label+'.xlsm');shutil.copyfile(TEMPLATE,path)
            case=dict(name=label,incoming_screen=(index%2==1),incoming_cancel_key=0 if index%2==0 else 2,incoming_water_suspend=(index%2==1));rec['cases'].append(case);save()
            with U.Session(rec,save,label) as s:
                s.book=s.app.Workbooks.Open(str(path),UpdateLinks=0)
                for source in sources:
                    comp=s.book.VBProject.VBComponents(source.stem)
                    if not a.negative:
                        comp.CodeModule.DeleteLines(1,comp.CodeModule.CountOfLines);comp.CodeModule.AddFromString(U.code(source))
                    assert U.code_key(comp.CodeModule.Lines(1,comp.CodeModule.CountOfLines))==U.code_key(U.code(source)),source.name
                    comp=None
                comp=s.book.VBProject.VBComponents('ThisWorkbook')
                if not a.negative:
                    comp.CodeModule.DeleteLines(1,comp.CodeModule.CountOfLines);comp.CodeModule.AddFromString(U.code(doc))
                assert U.code_key(comp.CodeModule.Lines(1,comp.CodeModule.CountOfLines))==U.code_key(U.code(doc));comp=None
                for name,text in [('WaterSourceFix3Probe',U.code(PROBE)),('Fix4StateProbe',HELPER)]:
                    comp=s.book.VBProject.VBComponents.Add(1);comp.Name=name;comp.CodeModule.AddFromString(text);comp=None
                s.compile();s.app.EnableEvents=True
                case['healthy_public_water']=list(s.call('WaterSourcePublicRun',2,False));save();assert case['healthy_public_water'][0]=='PASS'
                inp=s.book.Worksheets('浸透入力');water=s.book.Worksheets('浸透結果');result=s.book.Worksheets('解析結果')
                if label=='protected_output_coupling_formula':inp.Range('D4').Formula="=IF('浸透結果'!K2,1,0)";inp.Calculate()
                if label=='protected_output_dry':inp.Range('D4').Value2=0
                s.call('SeedSyntheticMain',label!='protected_output_dry');case['before']=observe(s)
                if label=='secondary_protected_soil_primary_gamma':
                    s.app.EnableEvents=False;inp.Range('B4').Value2=0;s.app.EnableEvents=True;result.Protect('fix4-owned-test')
                elif label!='success_coupled':
                    if label=='protected_output_coupling_formula':water.Range('K2').Locked=False;water.Range('L2').Locked=False
                    water.Protect('fix4-owned-test')
                response=list(s.call('RunWithIncomingState',case['incoming_screen'],case['incoming_cancel_key'],case['incoming_water_suspend']))
                case['run_message'],case['restored_screen'],case['restored_cancel_key'],case['restored_water_suspend']=response
                case['after']=observe(s);save()
                if label=='success_coupled':assert not case['run_message'] and case['after']['water_k2'] is True,case
                else:
                    assert '失敗しました' in case['run_message'],case
                    assert case['after']['journal']=='PENDING',case
                if label=='protected_output_coupling_formula':assert case['before']['coupling']==1 and case['after']['coupling']==0 and case['after']['water_k2'] is False,case
                if a.negative:
                    assert case['after']['flags'][0] and case['after']['pair']==[1.11,2.22] and case['after']['status']==case['before']['status'],case
                    case['status']='CONFIRMED_NEGATIVE'
                else:
                    assert case['restored_screen']==case['incoming_screen'] and case['restored_cancel_key']==case['incoming_cancel_key'] and case['restored_water_suspend']==case['incoming_water_suspend'],case
                    assert not case['after']['flags'][2] and not case['after']['flags'][3],case
                    if label=='protected_output_dry':assert case['after']['flags'][0] and case['after']['pair']==[1.11,2.22],case
                    else:
                        assert not case['after']['flags'][0] and not case['after']['flags'][1],case
                        if label=='secondary_protected_soil_primary_gamma':assert 'INVALID_WATER_UNIT_WEIGHT' in case['run_message'],case
                        else:assert case['after']['pair']==[None,None],case
                    case['status']='PASS'
                water.Unprotect('fix4-owned-test');result.Unprotect('fix4-owned-test');s.book.Save();water=result=inp=None;save()
        rec['status']='CONFIRMED_NEGATIVE' if a.negative else 'PASS'
    except BaseException as e:rec.update(status='FAIL',error=repr(e),traceback=traceback.format_exc());save();raise
    finally:
        rec['sha256_after']={str(x):U.sha(x) for x in monitored};rec['monitored_files_unchanged']=before==rec['sha256_after'];save()
    if not rec['monitored_files_unchanged'] or any(not s.get('owned_excel_exited') or s.get('dialogs') or s.get('monitor_errors') or s.get('hard_timeout') for s in rec['sessions']):
        rec.update(status='FAIL',post_run_verification_error='Ownership, immutability, dialogs, monitor or timeout failed');save()
        raise AssertionError(rec['post_run_verification_error'])
    print(json.dumps(dict(status=rec['status'],output=str(out),cases=len(rec['cases']),all_owned_excel_exited=True),ensure_ascii=False),flush=True)
if __name__=='__main__':main()
