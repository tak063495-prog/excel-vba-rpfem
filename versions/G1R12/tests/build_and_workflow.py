from pathlib import Path
import sys, json, shutil, hashlib, time, re
import win32com.client
from native_regression import ROOT, BASE, CHANGED, replace, setting
from native_helpers import compile_book
OUTPUT='RPFEM_20261009_G1R12_P06N_IssueFix.xlsm'
def normalize(text):
    # VBE capitalizes identifiers across the project; strings/comments remain exact.
    rows=[]
    for line in text.replace('\r','').splitlines():
        if line.startswith('Attribute VB_') or not line.strip():continue
        parts=re.split(r'("(?:[^"]|"")*"|\x27.*)',line.rstrip())
        rows.append(''.join(p if i%2 else p.casefold() for i,p in enumerate(parts)))
    return '\n'.join(rows)
def input_snapshot(book):
    return {s:book.Worksheets(s).UsedRange.Value for s in ['要素定義','材料データ','設定']}
def build():
    source=Path(BASE['path']); assert hashlib.sha256(source.read_bytes()).hexdigest()==BASE['sha256']
    target=ROOT/OUTPUT; assert not target.exists()
    shutil.copy2(source,target)
    app=win32com.client.DispatchEx('Excel.Application'); app.Visible=False; app.DisplayAlerts=False; app.EnableEvents=False
    book=None
    result={'scope':'production copy, native integration/compile/save/reopen; full P06N not run'}
    try:
        book=app.Workbooks.Open(str(target),UpdateLinks=0)
        before=input_snapshot(book)
        for name in CHANGED: replace(book,name,(ROOT/'src'/f'{name}.bas').read_text(encoding='utf-8'))
        compile_book(app,book)
        for p in (ROOT/'src').glob('*.bas'):
            actual=book.VBProject.VBComponents(p.stem).CodeModule
            (ROOT/'results'/('export_'+p.name)).write_text(actual.Lines(1,actual.CountOfLines).replace('\r',''),encoding='utf-8',newline='\n')
            assert normalize(actual.Lines(1,actual.CountOfLines))==normalize(p.read_text(encoding='utf-8')),p.name
        assert input_snapshot(book)==before
        guide=book.Worksheets('P06N 実行案内')
        guide.Range('A23').Value='G1R12: Issue #1の上界メッシュと場の一致、Issue #2のpivot付き因子分解失敗に対するHSD救済を修正。'
        guide.Range('A24').Value='小規模Fs根と出力・失敗注入は検証済み。G1R12の全P06N・25ケース全回帰は未実行。'
        book.Save(); book.Close(SaveChanges=False); book=None
        book=app.Workbooks.Open(str(target),UpdateLinks=0)
        compile_book(app,book)
        assert input_snapshot(book)==before
        for p in (ROOT/'src').glob('*.bas'):
            actual=book.VBProject.VBComponents(p.stem).CodeModule
            assert normalize(actual.Lines(1,actual.CountOfLines))==normalize(p.read_text(encoding='utf-8')),p.name
        assert len([book.VBProject.VBComponents(i).Name for i in range(1,book.VBProject.VBComponents.Count+1) if 'Issue' in book.VBProject.VBComponents(i).Name or 'Probe' in book.VBProject.VBComponents(i).Name])==0
        result.update({'pass':True,'compile':'PASS','reopen_compile':'PASS','sources_matched':46,'input_values_unchanged':True,'test_probes_absent':True,'sha256':hashlib.sha256(target.read_bytes()).hexdigest()})
    finally:
        (ROOT/'results/production_workbook.json').write_text(json.dumps(result,ensure_ascii=False,indent=2),encoding='utf-8')
        if book is not None: book.Close(SaveChanges=False)
        app.Quit()
    print('Production workbook saved/reopened/compiled/source match PASS',flush=True)
def workflows():
    rows=[]
    for version in ['G1R11','G1R12']:
        app=win32com.client.DispatchEx('Excel.Application'); app.Visible=False; app.DisplayAlerts=False; app.EnableEvents=False
        path=ROOT/'tests'/f'workflow_{version}.xlsm'
        shutil.copy2(BASE['path'] if version=='G1R11' else ROOT/OUTPUT,path)
        book=None
        try:
            book=app.Workbooks.Open(str(path),UpdateLinks=0)
            call=lambda name,*args:app.Run("'"+book.Name+"'!"+name,*args)
            book.Worksheets('要素定義').Range('A7:T1006').ClearContents()
            book.Worksheets('材料データ').Range('A5:Q1004').ClearContents()
            call('RPX_WriteSlopeInputs')
            book.Worksheets('材料データ').Range('A5:E5').Value=((1,5,45,18.82,45),)
            for key,val in [('TARGET',16),('ADAPT',1),('CYCLES',0),('QUADRATURE',2),('FS_TOL',.0001),('ANALYSIS_POLICY','ASSOCIATED'),('DIAGNOSTIC_MODE',1),('HSD_CAPTURE',0),('TEST_INJECTION',0)]: setting(book,key,val)
            compile_book(app,book)
            started=time.perf_counter(); call('RPX_Run'); seconds=time.perf_counter()-started
            ws=book.Worksheets('解析結果')
            row={'version':version,'scope':'actual RPX_Run adaptive, C5 phi=psi=45, TARGET16/CYCLES0/q2','elapsed_seconds':seconds,'status':ws.Range('A1').Value,'bounds':list(ws.Range('B4:C4').Value[0]),'elements':list(ws.Range('B10:C10').Value[0]),'upper_rows':book.Worksheets('要素データ').Range('A2:F1001').Value}
            assert all(isinstance(x,(int,float)) and x>0 for x in row['bounds']),row
            assert row['bounds'][0]<=row['bounds'][1],row
            rows.append(row)
            book.Save()
            print(version,'RPX_Run',row['status'],row['bounds'],row['elements'],round(seconds,3),'seconds',flush=True)
        finally:
            if book is not None:book.Close(SaveChanges=False)
            app.Quit()
    # Normal numerical path must retain roots and chosen meshes exactly.
    assert rows[0]['bounds']==rows[1]['bounds'],[(r['version'],r['bounds']) for r in rows]
    assert rows[0]['upper_rows']==rows[1]['upper_rows']
    for r in rows: r.pop('upper_rows')
    (ROOT/'results/native_workflow.json').write_text(json.dumps({'pass':True,'results':rows,'bounds_and_upper_tables_identical':True},ensure_ascii=False,indent=2),encoding='utf-8')
if __name__=='__main__':
    if 'build' in sys.argv:build()
    if 'workflow' in sys.argv:workflows()
