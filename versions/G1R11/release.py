from native_common import *
import re
def sha(p):return hashlib.sha256(p.read_bytes()).hexdigest()
def modules(b):return {c.Name:c.CodeModule.Lines(1,c.CodeModule.CountOfLines) if c.CodeModule.CountOfLines else '' for c in b.VBProject.VBComponents}
def semantic(s):
    s='\n'.join(s.replace('\r\r\n','\n').replace('\r\n','\n').splitlines()).strip()
    return re.sub(r'"(?:[^"\n]|"")*"|\x27[^\n]*|[A-Za-z_][A-Za-z_0-9]*',lambda m:m.group(0) if m.group(0).startswith(('"',"'")) else m.group(0).lower(),s)
for f in ['native_geometry.json','native_controlled.json','native_brackets.json','static_invariants.json']:
    assert (R/'results'/f).exists()
dest=R/'RPFEM_20261007_G1R11_P06N_Staged_1000.xlsm'
if dest.exists():
    prior=json.loads((R/'Release_Verification.json').read_text(encoding='utf-8'))
    assert sha(dest)==prior['sha256'], 'Preserve a file changed by the user'
    shutil.copy2(dest,R/'tests/Draft_before_staged_search_scope.xlsm')
    dest.unlink()
shutil.copy2(source,dest)
app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;book=None
try:
    book=app.Workbooks.Open(str(dest),UpdateLinks=0);expected=modules(book)
    before=book.Worksheets('設定').Range('A4:D100').Value2
    materials=book.Worksheets('材料データ').Range('A5:K500').Value2
    geometry=book.Worksheets('要素定義').Range('A7:T1000').Value2
    inject(book,False)
    for name in NAMES:expected[name]='\r\n'.join((R/'src'/(name+'.bas')).read_text(encoding='utf-8').splitlines()[1:])
    app.EnableEvents=False
    setvalue(book,'ADAPT_MESH_POLICY','LEGACY_THEN_REFINE');setvalue(book,'ADAPT_REFINE_CAP',1000);setvalue(book,'FS_BRACKET_POLICY','AUDITED_UPPER_LIMIT')
    ws=book.Worksheets('設定')
    ws.Cells(48,3).Value2='従来の再配置・細分化の後に、予算内で赤緑細分化を最大1回追加。LEGACYで従来メッシュ方式。'
    ws.Cells(49,3).Value2='追加細分化の上限。TARGETとの小さい方を使用。既存メッシュを縮小する設定ではありません。'
    ws.Cells(50,3).Value2='従来の前段800要素以下では探索を維持。追加細分化後の多層で監査済み上界を目安に使用し、下界の符号を実求解。単層・剛体は従来。'
    for row,choices in [(48,'LEGACY,ROBUST_REFINE,LEGACY_THEN_REFINE'),(50,'LEGACY,AUDITED_UPPER_LIMIT')]:
        ws.Cells(row,2).Validation.Delete();ws.Cells(row,2).Validation.Add(3,1,1,choices)
    ws.Cells(49,2).Validation.Delete();ws.Cells(49,2).Validation.Add(1,1,1,16,50000)
    after=ws.Range('A4:D100').Value2
    assert all(a==b for row,(a,b) in enumerate(zip(before,after),start=4) if row not in [48,49,50])
    book.Worksheets('解析結果').Cells.ClearContents()
    book.Worksheets('解析結果').Range('A1').Value2='G1R11 従来メッシュを土台とする改訂版・P06N設定済み。このブックの全Fs解析は未実行。'
    book.Worksheets('要素定義').Range('A3').Value2='P06N・Davis等価関連モデル・G1R11。従来の再配置→従来細分化→最大1000要素への追加細分化。G1R9 DD高速化・精度条件を維持。'
    for guide in book.Worksheets:
        if '実行案内' in guide.Name:
            guide.Range('A1').Value2='P06N G1R11 従来メッシュを土台とする改訂版 実行案内'
            guide.Range('B5').Value2='解析を実行し、G1R10の12時間21分・GAP14.85%と比較してください。全解析の短縮率・GAP1%到達は未検証です。'
            guide.Range('A17:B21').Value2=(('変更','従来の再配置と細分化を残し、追加細分化を最大1回・1000要素に制限。1536要素への直接成長を既定経路から外す。'),('Fs試行','多層の監査済み上界を目安にし、下界の符号を実求解。数値失敗時は従来探索へ復帰。単層・剛体は従来。'),('検証','PythonとExcelでメッシュ・固定Fs・監査・制御経路を確認。全P06N適応解析の完走検証は別途必要。'),('実行後','adapt_mesh_policy、adapt_action、adapt_staged_end、fs_upper_limit、audit、adapt_endを確認。'),('従来へ戻す','設定B48とB50をLEGACYに変更。追加細分化と新しいFs試行範囲制御を停止。'))
    assert book.Worksheets('材料データ').Range('A5:K500').Value2==materials
    assert book.Worksheets('要素定義').Range('A7:T1000').Value2==geometry
    compile_book(app,book);actual=modules(book)
    assert set(actual)==set(expected) and all(semantic(v)==semantic(actual[k]) for k,v in expected.items())
    assert all('ProbeBracket(' not in v and 'ProbeAdapt(' not in v and 'SmallRoots(' not in v and 'mockLower' not in v for v in actual.values())
    app.EnableEvents=True;book.Worksheets('要素定義').Activate();app.Goto(book.Worksheets('要素定義').Range('A1'),True)
    book.Save();book.Close(False);book=None;saved=sha(dest)
    book=app.Workbooks.Open(str(dest),UpdateLinks=0,ReadOnly=True);compile_book(app,book);actual=modules(book)
    assert set(actual)==set(expected) and all(semantic(v)==semantic(actual[k]) for k,v in expected.items())
    assert book.Worksheets('設定').Range('A4:D100').Value2==after
    assert book.Worksheets('材料データ').Range('A5:K500').Value2==materials
    assert book.Worksheets('要素定義').Range('A7:T1000').Value2==geometry
    settings={str(k):v for _,v,_,k in after if k}
    assert settings['TARGET']==1536 and settings['GAP']==.01 and settings['FS_TOL']==.0001 and settings['QUADRATURE']==8
    record={'workbook':str(dest),'sha256':saved,'source_sha256':sha(source),'changed_modules':NAMES,'native_compile':'PASS','saved_readback':'PASS','other_modules_and_document_events_unchanged':True,'no_test_or_mock_code_in_workbook':True,'materials_geometry_and_existing_settings_preserved':True,'G1R9_DD_speedup_preserved':True,'settings':settings,'full_P06N_adapt':'NOT_RUN','25_non_pore_cases':'NOT_RUN'}
    (R/'Release_Verification.json').write_text(json.dumps(record,ensure_ascii=False,indent=2),encoding='utf-8');print(record,flush=True)
finally:
    if book is not None:book.Close(False)
    app.Quit()
assert sha(source)==identity['sha256'] and sha(dest)==saved
