"""Owned Excel source-revision migration and pure public saved-book API regression."""
from pathlib import Path
from fix4_artifact_binding import bind
import argparse, base64, hashlib, importlib.util, json, shutil, traceback, xml.etree.ElementTree as ET
ROOT=Path(__file__).resolve().parents[1]; OLD=ROOT.parent/'G1R20_Fix3'
SPEC=importlib.util.spec_from_file_location('revision_session',OLD/'tests/run_saved_formula_excel.py')
U=importlib.util.module_from_spec(SPEC);SPEC.loader.exec_module(U)
NS='urn:rpfem:main-source:v1'

def physical(book):
    rows=[]
    for sheet,address in [('要素定義','A7:T1006'),('材料データ','A5:Q1004'),('設定','A4:D120')]:
        cells=book.Worksheets(sheet).Range(address);rows.append((sheet,address,cells.Value2,cells.Formula));cells=None
    return hashlib.sha256(json.dumps(rows,ensure_ascii=False).encode('utf8')).hexdigest()

def proof(book):
    parts=book.CustomXMLParts.SelectByNamespace(NS)
    if parts.Count!=1:return dict(count=parts.Count)
    xml=parts.Item(1).XML; node=ET.fromstring(xml)
    return dict(count=1,identifier=base64.b64decode(node.find('{'+NS+'}identifier').text).decode('utf-16-le'),xml_sha256=hashlib.sha256(xml.encode('utf8')).hexdigest())

def main():
    p=argparse.ArgumentParser();p.add_argument('--execute',action='store_true');p.add_argument('--task',choices=['migration','pure_public'],required=True);p.add_argument('--output',type=Path,required=True);a=p.parse_args()
    if not a.execute:raise SystemExit('Requires coordinated --execute')
    out=a.output.resolve();out.mkdir(parents=True,exist_ok=False)
    template=(OLD if a.task=='migration' else ROOT)/'delivery'/('RPFEM_20261017_G1R20_Fix3.xlsm' if a.task=='migration' else 'RPFEM_20261017_G1R20_Fix4.xlsm')
    path=out/(a.task+'.xlsm');shutil.copyfile(template,path)
    sources=sorted((ROOT/'src/vba').glob('*.bas'));doc=ROOT/'src/workbook/ThisWorkbook.code.txt'
    monitored=[template,Path(__file__),*sources,doc,ROOT/'delivery/rpfem_native_x64.dll'];monitored.append(ROOT/'tests/fix4_artifact_binding.py');before={str(x):U.sha(x) for x in monitored}
    rec=dict(status='PREPARING',task=a.task,sessions=[],sha256_before=before,calls=[])
    rec['final_artifact_binding']=bind(ROOT)
    def save():(out/'summary.json').write_text(json.dumps(rec,ensure_ascii=False,indent=2)+'\n',encoding='utf8')
    save();shutil.copyfile(__file__,out/'runner_snapshot.py')
    try:
        if a.task=='migration':
            rec['fixture']='Synthetic saved display/source record only; no numerical certificate or conic optimizer call'
            with U.Session(rec,save,'prepare_fix3_saved_proof') as s:
                s.book=s.app.Workbooks.Open(str(path),UpdateLinks=0);s.call('RPX_Setup')
                result=s.book.Worksheets('解析結果');result.Range('A1').Value2='SYNTHETIC_FIX3_SAVED_PROOF';result.Range('B4:C4').Value2=((1.11,2.22),);result=None
                s.call('RPX_InputObserve');s.call('RPX_InputCommit')
                rec['old_proof']=proof(s.book);rec['physical_before']=physical(s.book)
                assert rec['old_proof']['identifier']=='RPX_MAIN_SOURCE_V1'
                s.book.Save();save()
            with U.Session(rec,save,'upgrade_sources_without_reanalysis') as s:
                s.book=s.app.Workbooks.Open(str(path),UpdateLinks=0)
                for source in sources:
                    comp=s.book.VBProject.VBComponents(source.stem);comp.CodeModule.DeleteLines(1,comp.CodeModule.CountOfLines);comp.CodeModule.AddFromString(U.code(source))
                    assert U.code_key(comp.CodeModule.Lines(1,comp.CodeModule.CountOfLines))==U.code_key(U.code(source));comp=None
                comp=s.book.VBProject.VBComponents('ThisWorkbook');comp.CodeModule.DeleteLines(1,comp.CodeModule.CountOfLines);comp.CodeModule.AddFromString(U.code(doc));comp=None;s.compile()
                rec['proof_after_upgrade']=proof(s.book);rec['physical_after_upgrade']=physical(s.book)
                assert rec['old_proof']==rec['proof_after_upgrade'] and rec['physical_before']==rec['physical_after_upgrade']
                s.book.Save();save()
            with U.Session(rec,save,'reopen_fix4_events_enabled') as s:
                s.app.EnableEvents=True;s.book=s.app.Workbooks.Open(str(path),UpdateLinks=0)
                checks={source.name:U.code_key(s.book.VBProject.VBComponents(source.stem).CodeModule.Lines(1,s.book.VBProject.VBComponents(source.stem).CodeModule.CountOfLines))==U.code_key(U.code(source)) for source in sources}
                checks['ThisWorkbook']=U.code_key(s.book.VBProject.VBComponents('ThisWorkbook').CodeModule.Lines(1,s.book.VBProject.VBComponents('ThisWorkbook').CodeModule.CountOfLines))==U.code_key(U.code(doc))
                rec['reopened_full_64_code_comparison']=checks;assert all(checks.values())
                result=s.book.Worksheets('解析結果');rec['reopened']=dict(pair=list(result.Range('B4:C4').Value2[0]),status=result.Range('A1').Value2);result=None
                assert rec['reopened']['pair']==[None,None] and '入力変更' in rec['reopened']['status'],rec['reopened']
                rec['physical_after_reopen']=physical(s.book);assert rec['physical_after_reopen']==rec['physical_before']
                s.book.Save();save()
        else:
            rec['fixture']='Final saved Fix4 workbook; no VBE access, no injected modules or replacement numerical state; direct public macros only'
            shutil.copyfile(ROOT/'delivery/rpfem_native_x64.dll',out/'rpfem_native_x64.dll');shutil.copyfile(OLD/'tests/read_model_shear_c10.txt',out/'input.txt')
            with U.Session(rec,save,'pure_public_saved_fix4') as s:
                s.book=s.app.Workbooks.Open(str(path),UpdateLinks=0)
                for cycle in range(3):
                    for name,args in [('RPX_TestFs',(str(out/'input.txt'),str(out/('fs_'+str(cycle)+'.csv')))),('RPX_Setup',()),('RPX_SeismicSetup',())]:
                        rec['calls'].append(dict(cycle=cycle,function=name,status='CALLING'));save()
                        answer=s.call(name,*args);rec['calls'][-1].update(status='PASS',answer=answer);save()
                        if name=='RPX_TestFs':assert str(answer).startswith('PASS'),answer
                rec['native_info']=s.call('RPX_NativeInfo');s.book.Save();save()
        rec['status']='PASS'
    except BaseException as error:rec.update(status='FAIL',error=repr(error),traceback=traceback.format_exc());save();raise
    finally:
        rec['sha256_after']={str(x):U.sha(x) for x in monitored};rec['monitored_files_unchanged']=before==rec['sha256_after'];save()
    if not rec['monitored_files_unchanged'] or any(not x.get('owned_excel_exited') or x.get('dialogs') or x.get('monitor_errors') or x.get('hard_timeout') for x in rec['sessions']):
        rec.update(status='FAIL',post_run_verification_error='Ownership, immutability, dialogs, monitor or timeout failed');save()
        raise AssertionError(rec['post_run_verification_error'])
    print(json.dumps(dict(status=rec['status'],task=a.task,sessions=len(rec['sessions']),out=str(out)),ensure_ascii=False),flush=True)
if __name__=='__main__':main()
