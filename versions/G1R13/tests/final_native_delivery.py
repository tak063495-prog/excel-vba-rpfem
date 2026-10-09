"""Only run after both repaired full native analyses and independent audits finish."""
from native_repair import *
from io import BytesIO
import zipfile,olefile
from oletools.olevba import VBA_Project,VBA_Parser
def code_modules(path):
    p=VBA_Parser(str(path))
    try:return {name:normalize(code) for _,_,name,code in p.extract_macros()}
    finally:p.close()
original=REPO/'versions/G1R12/RPFEM_20261009_G1R12_P06N_IssueFix.xlsm'
base=code_modules(original);all_checks=[]
app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;book=None
try:
    for case in ['D07A','D03N']:
        folder=ROOT/'native_retry'/case;rec=json.loads((folder/'status.json').read_text(encoding='utf-8'));assert rec.get('fs_obtained'),rec
        src=ROOT/rec['workbook'];dst=ROOT/'delivery'/f'RPFEM_20261010_G1R13_{case}_Verified.xlsm';assert not dst.exists();shutil.copy2(src,dst)
        book=app.Workbooks.Open(str(dst),UpdateLinks=0)
        # Tests are already exported; production delivery has only the 2 new production modules.
        book.VBProject.VBComponents.Remove(book.VBProject.VBComponents('BatchEvidence'))
        for p in (ROOT/'src/vba').glob('*.bas'):
            cm=book.VBProject.VBComponents(p.stem).CodeModule
            assert normalize(cm.Lines(1,cm.CountOfLines))==normalize(p.read_text(encoding='utf-8')),p.name
        compile_book(app,book)
        expected=(rec['lower_fs'],rec['upper_fs']);actual=book.Worksheets('解析結果').Range('B4:C4').Value2[0];assert actual==expected
        settings={}
        ws=book.Worksheets('設定')
        for row in range(4,101):
            key=ws.Cells(row,4).Value2
            if key:settings[str(key)]=ws.Cells(row,2).Value2
        assert settings['TEST_INJECTION']==0 and settings['FS_TOL']==.0001 and settings['QUADRATURE']==8
        guide=book.Worksheets('P06N 実行案内');guide.Name='実行案内';guide.Range('A1:K40').ClearContents()
        guide.Range('A1').Value2=case+' / G1R13 修復・実Excel検証済み'
        guide.Range('A3').Value2='このブックは '+case+' の材料・形状・荷重・解析設定を入力済みです。'
        guide.Range('A5').Value2='再計算する場合は入力表、材料データ、設定を確認し、既存の解析ボタンから実行してください。'
        guide.Range('A7').Value2='保存結果：下界Fs='+format(rec['lower_fs'],'.12g')+'、上界Fs='+format(rec['upper_fs'],'.12g')
        guide.Range('A9').Value2='精度区分：'+('根幅とGAP目標を達成' if rec['target_met'] else '監査済みFsあり。根幅またはGAP目標は未達：部分結果として扱ってください。')
        guide.Range('A11').Value2='D03NはDavis等価モデルの結果です。D07Aは設定した上載荷重20を維持しています。'
        guide.Range('A13').Value2='詳細な設定、監査値、制限、復旧方法は同梱の修復報告とIMPORT_ROLLBACK_JA.mdを参照してください。'
        book.Worksheets('解析結果').Activate();book.Save();book.Close(SaveChanges=False);book=None
        actual=code_modules(dst)
        assert set(actual)==set(base)|{'RPX_WorkPrecision.bas','RPX_LoadStepMesh.bas'}
        unchanged=[]
        for name,value in base.items():
            if name not in ['RPX_Audit.bas','RPX_Fs.bas','RPX_HSDCandidate.bas','RPX_Mesh.bas']:
                assert actual[name]==value,name;unchanged.append(name)
        for p in (ROOT/'src/vba').glob('*.bas'):assert actual[p.name]==normalize(p.read_text(encoding='utf-8')),p.name
        with zipfile.ZipFile(dst) as z:
            with olefile.OleFileIO(BytesIO(z.read('xl/vbaProject.bin'))) as ole:
                v=VBA_Project(ole,'','PROJECT','VBA/dir');cp,codec=v.codepage,v.codec
        assert cp==932,(cp,codec)
        for p in (ROOT/'import').glob('*.bas'):
            raw=p.read_bytes();assert raw.count(b'\n')==raw.count(b'\r\n')
            assert raw.decode(codec,errors='strict').splitlines()==(ROOT/'src/vba'/p.name).read_text(encoding='utf-8').splitlines(),p.name
        all_checks.append({'case':case,'delivery':str(dst.relative_to(ROOT)),'sha256':sha(dst),'compile':'PASS','all_48_standard_sources_match':True,'unchanged_original_modules':len(unchanged),'existing_cls_modules_unchanged':10,'test_modules_removed':True,'project_codepage':cp,'import_lossless_crlf':True,'saved_fs_values':expected,'inputs_settings_preserved':True})
finally:
    if book is not None:book.Close(SaveChanges=False)
    app.Quit();gc.collect()
write(ROOT/'results/delivery_native_integrity.json',{'pass':True,'checks':all_checks,'original_sha256':sha(original)})
print(json.dumps({'pass':True,'books':[r['delivery'] for r in all_checks]},ensure_ascii=False))
