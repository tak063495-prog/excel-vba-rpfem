"""Owned serial Excel GenerateMesh-only exports; no imported field and no solve."""
from pathlib import Path
import argparse, gc, hashlib, importlib.util, json, shutil, sys, threading, time, uuid
import numpy as np
import psutil, win32com.client, win32con, win32gui, win32process

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('owned_harness',ROOT/'tests/run_adaptive_excel.py')
h=importlib.util.module_from_spec(spec);spec.loader.exec_module(h)


def read_geometry(path):
    raw=path.read_bytes();pos=8
    if raw[:8]!=b'RP25G001':raise ValueError('Not geometry-only format')
    def take(dtype,n):
        nonlocal pos
        a=np.frombuffer(raw,dtype=dtype,count=n,offset=pos).copy();pos+=a.nbytes;return a
    nn,ne,nm,nr,q=map(int,take('<i4',5));scale=float(take('<f8',1)[0])
    if min(nn,ne,nm,q)<=0 or nr!=0:raise ValueError('Unsupported geometry')
    xy=take('<f8',2*nn).reshape(nn,2);cells=take('<i4',5*ne).reshape(ne,5)
    mats=take('<f8',4*nm).reshape(nm,4);body=take('<f8',4*ne).reshape(ne,2,2)
    bd={};kinds=['load','fixed','roller_x','roller_y']
    for _ in range(int(take('<i4',1)[0])):
        a,b,k,r=map(int,take('<i4',4));edge=tuple(sorted((a,b)))
        if edge in bd:raise ValueError('Duplicate boundary')
        bd[edge]=dict(kind=kinds[k],rigid=r,loads=take('<f8',4).reshape(2,2).T)
    if pos!=len(raw) or np.any(cells[:,4]!=-1):raise ValueError('Trailing bytes or rigid')
    return xy,cells,mats,body,bd,dict(q=q,stress_scale=scale)


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--preset-only',action='store_true');args=parser.parse_args()
    prior=[p.pid for p in psutil.process_iter() if p.name().upper()=='EXCEL.EXE']
    if prior:raise RuntimeError('Another Excel exists: '+repr(prior))
    out=ROOT/'results'/('mesh_only_'+str(time.time_ns()));out.mkdir()
    bookpath=out/'MeshOnly.xlsm';shutil.copyfile(h.BASE/'delivery/RPFEM_20261015_G1R19.xlsm',bookpath)
    shutil.copyfile(h.BASE/'delivery/rpfem_native_x64.dll',out/'rpfem_native_x64.dll')
    rec=dict(status='PREPARING',run_token=uuid.uuid4().hex,prior_excel_pids=prior,
             source_sha256={n:h.sha(p) for n,p in h.sources().items()},
             dll_sha256=h.sha(out/'rpfem_native_x64.dll'),solver_calls=0,field_or_root_imported=False,
             harness_sha256=h.sha(__file__),probe_sha256=h.sha(ROOT/'tests/MeshOnlyProbe.bas'),
             published_workbook_sha256=h.sha(h.BASE/'delivery/RPFEM_20261015_G1R19.xlsm'),
             units='Canonical benchmark native coordinates and physical material/load units unchanged',
             coarse_target_rule='CLng(1000*0.35), bounded by [256,600] and final target => 350',
             cases={},dialogs=[])
    save=lambda:(out/'summary.json').write_text(json.dumps(rec,indent=2,ensure_ascii=False,allow_nan=False)+'\n',encoding='utf8')
    for p in [Path(__file__),ROOT/'tests/MeshOnlyProbe.bas']:shutil.copyfile(p,out/p.name)
    for n,p in h.sources().items():
        d=out/'source_snapshot'/p.name;d.parent.mkdir(exist_ok=True);shutil.copyfile(p,d)
    app=win32com.client.DispatchEx('Excel.Application');pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    rec.update(owned_excel_pid=pid,owned_excel_creation_time=psutil.Process(pid).create_time());save()
    stop=threading.Event()
    def watch():
        while not stop.wait(.25):
            def visit(w,_):
                if win32process.GetWindowThreadProcessId(w)[1]!=pid or win32gui.GetClassName(w)!='#32770':return
                children=[];win32gui.EnumChildWindows(w,lambda c,_:children.append((c,win32gui.GetWindowText(c))),None)
                text=[t for _,t in children]
                if text not in rec['dialogs']:rec['dialogs'].append(text)
                for c,t in children:
                    if t in ['OK','&OK'] or ('Visual Basic' in win32gui.GetWindowText(w) and ('終了' in t or '&End' in t)):win32gui.PostMessage(c,win32con.BM_CLICK,0,0)
            win32gui.EnumWindows(visit,None)
    watcher=threading.Thread(target=watch,daemon=True);watcher.start();book=c=control=None
    try:
        app.Visible=False;app.DisplayAlerts=False;app.EnableEvents=False;app.MultiThreadedCalculation.Enabled=False
        book=app.Workbooks.Open(str(bookpath),UpdateLinks=0)
        for n,p in h.sources().items():
            c=book.VBProject.VBComponents(n);c.CodeModule.DeleteLines(1,c.CodeModule.CountOfLines);c.CodeModule.AddFromString(h.vba(p))
        c=book.VBProject.VBComponents.Add(1);c.Name='MeshOnlyProbe';c.CodeModule.AddFromString(h.vba(ROOT/'tests/MeshOnlyProbe.bas'))
        c.CodeModule.CodePane.Show();control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled:control.Execute();time.sleep(.25)
        if control.Enabled:raise RuntimeError('VBA compile failed')
        app.VBE.MainWindow.Visible=False;rec['compile']='PASS';save()
        call=lambda name,*v:app.Run("'"+book.Name+"'!"+name,*v)
        if args.preset_only:
            rec['preset_unit']=list(call('MeshOnlyPresetUnit'));save()
            if rec['preset_unit'][0]!='PASS':raise RuntimeError(rec['preset_unit'])
        for case in ([] if args.preset_only else ['P06A','P06N']):
            inp=h.REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs'/(case+'.json')
            data=json.loads(inp.read_text(encoding='utf8'));folder=out/case;folder.mkdir();shutil.copyfile(inp,folder/'input_record.json')
            call('RPX_Setup')
            for sheet,ranges in [('要素定義',[('A','C','points',7),('E','H','regions',7),('J','T','faces',7)]),('材料データ',[('A','E','materials',5)])]:
                ws=book.Worksheets(sheet);ws.Range('A7:T1006' if sheet=='要素定義' else 'A5:Q1004').ClearContents()
                for col,last,key,row in ranges:
                    vals=tuple(tuple(x) for x in data[key]);ws.Range(f'{col}{row}:{last}{row+len(vals)-1}').Value2=vals
                ws=None
            for k,v in data['settings'].items():call('MeshOnlySetting',k,v)
            call('RPX_Speed2pctPreset');book.Worksheets('浸透入力').Range('D4').Value2=0
            settings={str(r[2]):r[0] for r in book.Worksheets('設定').Range('B4:D999').Value2 if r[2]}
            rec['cases'][case]=dict(input_sha256=h.sha(inp),settings=settings,targets={})
            for target in [350,1000]:
                result=list(call('MeshOnlyExport',target,str(folder/f'target_{target}.bin'),str(folder/f'target_{target}.context.tsv')))
                rec['cases'][case]['targets'][str(target)]={'result':result}
                if result[0]!='PASS':raise RuntimeError(result)
                save()
        book.Save();rec['status']='EXPORTED';save()
    except BaseException as e:rec.update(status='FAIL',error=repr(e));save();raise
    finally:
        if book is not None:book.Close(False)
        app.Quit();stop.set();watcher.join(2);book=c=control=app=None;gc.collect()
        for _ in range(50):
            if not psutil.pid_exists(pid):break
            time.sleep(.1)
        rec['owned_excel_exited']=not psutil.pid_exists(pid);save()
    sys.path.insert(0,str(h.REPO/'validation/accuracy_20261013/python'))
    import accuracy_driver as driver
    import experiment_core as core
    for case,item in rec['cases'].items():
        _,model=driver.load_model(case)
        for target,entry in item['targets'].items():
            path=out/case/f'target_{target}.bin';xy,cells,mats,body,bd,meta=read_geometry(path)
            mesh=core.mesh_from_arrays(xy,cells[:,:3],cells[:,3],model)
            if not np.array_equal(mesh.materials,mats) or not np.array_equal(mesh.body,body) or mesh.stress_scale!=meta['stress_scale']:raise ValueError('Canonical physical model mismatch')
            if mesh.boundary.keys()!=bd.keys():raise ValueError('Boundary edge mismatch')
            for edge,want in mesh.boundary.items():
                got=bd[edge]
                if got['kind']!=want['kind'] or got['rigid']!=want['rigid'] or not np.array_equal(got['loads'],want['loads']):raise ValueError('Boundary value mismatch')
            dest=out/case/f'target_{target}_mesh.npz';driver.save_mesh(dest,mesh)
            entry.update(canonical_physics_exact=True,geometry_sha256=h.sha(path),mesh_npz_sha256=h.sha(dest),nodes=len(xy),elements=len(cells),q=meta['q'],field_present=False,root_present=False)
    rec['status']='PASS';rec['files']={str(p.relative_to(out)):{'bytes':p.stat().st_size,'sha256':h.sha(p)} for p in out.rglob('*') if p.is_file() and p.name!='summary.json'};save()
    print(json.dumps({'status':'PASS','output':str(out),'cases':rec['cases'],'owned_excel_exited':rec['owned_excel_exited']}))

if __name__=='__main__':main()
