from native_repair import *
folder=ROOT/'mesh25_retry';folder.mkdir(exist_ok=True)
source=REPO/'versions/G1R12/RPFEM_20261009_G1R12_P06N_IssueFix.xlsm'
current=ROOT/'cases/D07A/RPFEM_20261010_G1R13_D07A.xlsm'
app=win32com.client.DispatchEx('Excel.Application');app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False
records={};book=None
try:
    for tag,src in [('before',source),('after',current)]:
        path=folder/f'{tag}.xlsm';assert not path.exists();shutil.copy2(src,path)
        book=app.Workbooks.Open(str(path),UpdateLinks=0)
        inject(book,'MeshComparison',(ROOT/'tests/MeshComparison.bas').read_text(encoding='utf-8'));compile_book(app,book)
        for p in sorted((ROOT/'inputs').glob('*.json')):
            rec=json.loads(p.read_text(encoding='utf-8'));case=p.stem
            ws=book.Worksheets('要素定義');ws.Range('A7:T1006').ClearContents()
            for first,last,name in [('A','C','points'),('E','H','regions'),('J','T','faces')]:
                data=tuple(tuple(v) for v in rec[name]);ws.Range(first+'7:'+last+str(6+len(data))).Value2=data
            ws=book.Worksheets('材料データ');ws.Range('A5:Q1004').ClearContents();data=tuple(tuple(v) for v in rec['materials']);ws.Range('A5:E'+str(4+len(data))).Value2=data
            ws=book.Worksheets('設定')
            for row in range(4,101):
                key=ws.Cells(row,4).Value2
                if key in rec['settings']:ws.Cells(row,2).Value2=rec['settings'][key]
            vals=call(app,book,'CompareMesh',str(folder/f'{case}_{tag}.model.txt')).split('\t')
            assert vals[4]=='raw';assert abs(float(vals[2])-rec['source_case']['area'])<=1e-9*max(1.,rec['source_case']['area'])
            assert abs(float(vals[3])-rec['total_weight'])<=1e-9*max(1.,rec['total_weight'])
            records.setdefault(case,{})[tag]=vals
            print(tag,case,vals[:2],flush=True)
        book.Close(SaveChanges=False);book=None
    changed=[]
    for case,rec in records.items():
        a=(folder/f'{case}_before.model.txt').read_bytes();b=(folder/f'{case}_after.model.txt').read_bytes()
        rec['exact_same_physical_mesh_model']=a==b
        if a!=b:changed.append(case)
    assert changed==['D07A'],changed
    assert records['D07A']['before'][:2]==['314','553'] and records['D07A']['after'][:2]==['316','557']
    write(folder/'result.json',{'scope':'native coarse mesh construction only, no Fs or target PASS claims','cases':records,'changed_cases':changed,'unchanged_exact_models':24,'area_and_weight_all_pass':True,'pass':True})
finally:
    if book is not None:book.Close(SaveChanges=False)
    app.Quit();gc.collect()
