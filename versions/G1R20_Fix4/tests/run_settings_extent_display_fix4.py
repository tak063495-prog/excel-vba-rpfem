"""Final-book display invalidation for an unsupported settings key; synthetic only."""
from pathlib import Path
from fix4_artifact_binding import bind
import argparse, importlib.util, json, shutil, traceback
ROOT=Path(__file__).resolve().parents[1];OLD=ROOT.parent/'G1R20_Fix3'
SPEC=importlib.util.spec_from_file_location('extent_session',OLD/'tests/run_saved_formula_excel.py')
U=importlib.util.module_from_spec(SPEC);SPEC.loader.exec_module(U)

def main():
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--negative',action='store_true');p.add_argument('--output',required=True,type=Path);a=p.parse_args()
    if not a.execute:raise SystemExit('Requires coordinated --execute')
    out=a.output.resolve();out.mkdir(parents=True,exist_ok=False)
    template=ROOT/'delivery/RPFEM_20261017_G1R20_Fix4.xlsm'
    probe=ROOT/'tests/SettingsExtentFix4Probe.bas'
    monitored=[template,Path(__file__),*sorted((ROOT/'src/vba').glob('*.bas')),ROOT/'src/workbook/ThisWorkbook.code.txt',probe]
    monitored.append(ROOT/'tests/fix4_artifact_binding.py');before={str(x):U.sha(x) for x in monitored};snap=out/'source_snapshot';snap.mkdir()
    for item in monitored[2:]:shutil.copyfile(item,snap/item.name)
    shutil.copyfile(__file__,out/'runner_snapshot.py')
    rec=dict(status='PREPARING',negative=a.negative,sessions=[],cases={},sha256_before=before,fixture='Synthetic saved display and source record; final saved Fix4 book, display paths use no VBE; final public-reader error is caught in one test-only VBA probe; no numerical certificate or optimizer')
    rec['final_artifact_binding']=bind(ROOT)
    def save():(out/'summary.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    def display(s):
        result=s.book.Worksheets('解析結果');ws=s.book.Worksheets('設定')
        row=dict(pair=list(result.Range('B4:C4').Value2[0]),status=result.Range('A1').Value2,key=ws.Range('D121').Value2,value=ws.Range('B121').Value2);result=ws=None;return row
    def seed(s):
        s.call('RPX_Setup');result=s.book.Worksheets('解析結果');result.Range('A1').Value2='SYNTHETIC_MATCHING_FIX4_SAVED_SOURCE';result.Range('B4:C4').Value2=((1.11,2.22),);result=None
    save()
    try:
        path=out/'formula_tail.xlsm';shutil.copyfile(template,path)
        with U.Session(rec,save,'formula_tail') as s:
            s.book=s.app.Workbooks.Open(str(path),UpdateLinks=0);seed(s)
            s.call('RPX_InputObserve');s.call('RPX_InputCommit');s.app.EnableEvents=True
            s.book.Worksheets('設定').Range('D121').Value2='SEISMIC_KH'
            rec['direct_edit_control']=display(s);assert rec['direct_edit_control']['pair']==[None,None]
            s.app.EnableEvents=False;seed(s)
            ref=s.book.Worksheets.Add(After=s.book.Worksheets(s.book.Worksheets.Count));ref.Name='参照入力';ref.Range('A1').Value2=0
            ws=s.book.Worksheets('設定');ws.Range('B121').Value2=.2
            s.call('RPX_InputObserve');s.call('RPX_InputCommit')
            ws.Range('D121').Formula='=IF(\'参照入力\'!A1=1,"SEISMIC_KH","")';ws.Calculate()
            rec['cases']['formula_tail']=dict(before=display(s));save()
            s.app.EnableEvents=True;ref.Range('A1').Value2=1;s.app.CalculateFull()
            rec['cases']['formula_tail']['after']=display(s);ref=ws=None
            if a.negative:assert rec['cases']['formula_tail']['after']['pair']==[1.11,2.22]
            else:assert rec['cases']['formula_tail']['after']['pair']==[None,None]
            s.book.Save();save()
        path=out/'saved_tail.xlsm';shutil.copyfile(template,path)
        with U.Session(rec,save,'save_events_off_tail') as s:
            s.book=s.app.Workbooks.Open(str(path),UpdateLinks=0);seed(s)
            ws=s.book.Worksheets('設定');ws.Range('B121').Value2=.2;s.call('RPX_InputObserve');s.call('RPX_InputCommit')
            rec['cases']['saved_tail']=dict(before=display(s));ws.Range('D121').Value2='SEISMIC_KH';ws=None;s.book.Save();save()
        with U.Session(rec,save,'reopen_saved_tail') as s:
            s.app.EnableEvents=True;s.book=s.app.Workbooks.Open(str(path),UpdateLinks=0)
            rec['cases']['saved_tail']['after']=display(s)
            if a.negative:assert rec['cases']['saved_tail']['after']['pair']==[1.11,2.22]
            else:assert rec['cases']['saved_tail']['after']['pair']==[None,None]
            # Unhandled VBA errors can open a modal runtime-error dialog instead of
            # preserving Err.Description through Application.Run. Catch the public
            # reader's error inside a test-only module in this disposable copy.
            s.component=s.book.VBProject.VBComponents.Add(1);s.component.Name=probe.stem;s.component.CodeModule.AddFromString(U.code(probe))
            assert U.code_key(s.component.CodeModule.Lines(1,s.component.CodeModule.CountOfLines))==U.code_key(U.code(probe))
            comparison={source.name:U.code_key(s.book.VBProject.VBComponents(source.stem).CodeModule.Lines(1,s.book.VBProject.VBComponents(source.stem).CodeModule.CountOfLines))==U.code_key(U.code(source)) for source in sorted((ROOT/'src/vba').glob('*.bas'))}
            doc=ROOT/'src/workbook/ThisWorkbook.code.txt';comparison[doc.name]=U.code_key(s.book.VBProject.VBComponents('ThisWorkbook').CodeModule.Lines(1,s.book.VBProject.VBComponents('ThisWorkbook').CodeModule.CountOfLines))==U.code_key(U.code(doc))
            rec['public_reader_product_64_code_comparison']=comparison;assert len(comparison)==64 and all(comparison.values())
            s.compile();answer=s.call('ExtentReadInputsError');rec['public_reader_error']=answer
            if answer=='OK':
                rec['reader_accepted_unsupported']=True
                if not a.negative:raise AssertionError('Unsupported row was accepted by public input reader')
            else:
                number,description=answer.split('|',1);rec['read_inputs_rejected']=dict(number=int(number),description=description)
                assert int(number)==5 and 'SETTING_ROW_OUT_OF_RANGE' in description
            s.book.Save();save()
        rec['status']='CONFIRMED_NEGATIVE' if a.negative else 'PASS'
    except BaseException as error:rec.update(status='FAIL',error=repr(error),traceback=traceback.format_exc());save();raise
    finally:
        rec['sha256_after']={str(x):U.sha(x) for x in monitored};rec['monitored_files_unchanged']=before==rec['sha256_after'];save()
    if not rec['monitored_files_unchanged'] or any(not x.get('owned_excel_exited') or x.get('dialogs') or x.get('monitor_errors') or x.get('hard_timeout') for x in rec['sessions']):
        rec.update(status='FAIL',post_run_verification_error='Ownership, immutability, dialogs, monitor or timeout failed');save()
        raise AssertionError(rec['post_run_verification_error'])
    print(json.dumps(dict(status=rec['status'],cases=2,sessions=3,out=str(out)),ensure_ascii=False),flush=True)
if __name__=='__main__':main()
