"""Owned Excel original-program audit of Python fields, with no native solver run."""
from pathlib import Path
import argparse
import gc
import hashlib
import importlib.util
import json
import math
import shutil
import struct
import threading
import time
import psutil
import win32com.client
import win32con
import win32gui
import win32process
from export_python_witness import export, read_state, np, driver, core

ROOT = Path(__file__).resolve().parents[1]
REPO = ROOT.parents[1]
spec = importlib.util.spec_from_file_location('adapt_harness', Path(__file__).with_name('test_adapt_excel.py'))
adapt = importlib.util.module_from_spec(spec)
spec.loader.exec_module(adapt)


def replay_original(folder, mode, program_name='audited_original.bin', field_name='audited_field.bin'):
    """Read literal VBA raw rows/cones with an independent compensated sum."""
    specification=importlib.util.spec_from_file_location('original_reader',ROOT.parent/'G1R18_RayGate/tests/replay_original_guide.py')
    replay=importlib.util.module_from_spec(specification);specification.loader.exec_module(replay)
    fields=replay.Reader(folder/field_name,b'P06X0001')
    n=fields.read('i');x=[fields.read('d') for _ in range(n)];fields.finish()
    assert n>0 and all(math.isfinite(v) for v in x)
    reader=replay.Reader(folder/program_name,b'P06VBA01')
    n_original,m,blocks=reader.read('3i');assert n_original==n and m>0 and blocks>0
    cost=[reader.read('d') for _ in range(n)];assert all(math.isfinite(v) for v in cost)
    def residual():
        count,offset=reader.read('id');assert 0<=count<=n and math.isfinite(offset)
        terms=[offset]
        for _ in range(count):
            index,value=reader.read('id');assert 0<=index<n and math.isfinite(value)
            terms.append(value*x[index])
        return math.fsum(terms)
    raw_eq=0.
    for _ in range(m):
        count,rhs=reader.read('id');assert 0<=count<=n and math.isfinite(rhs)
        terms=[-rhs]
        for _ in range(count):
            index,value=reader.read('id');assert 0<=index<n and math.isfinite(value)
            terms.append(value*x[index])
        raw_eq=max(raw_eq,abs(math.fsum(terms)))
    raw_cone=0.
    for _ in range(blocks):
        dim=reader.read('i');assert dim in (1,3)
        values=[residual() for _ in range(dim)]
        raw_cone=max(raw_cone,-values[0] if dim==1 else math.hypot(*values[1:])-values[0])
    reader.finish();objective=math.fsum(c*v for c,v in zip(cost,x))
    assert raw_eq<=1e-7 and raw_cone<=1e-7
    assert (objective<=.99999901 if mode=='upper' else -objective>=1.000001)
    return dict(raw_equality=raw_eq,raw_cone=raw_cone,objective=objective,raw_original_rows=m,cones=blocks)


def run(case, mode, source):
    inputs_path = REPO / 'validation/accuracy_20261013/python/data' / (case + '.json')
    if not inputs_path.exists():
        inputs_path = driver.CATALOG / (case + '.json')
    inputs = json.loads(inputs_path.read_text(encoding='utf-8'))
    if inputs.get('water', {}).get('coupling', 0) != 0:
        raise ValueError('A coupled pore snapshot needs a separate native import audit')
    folder = ROOT / 'results' / ('python_field_excel_' + str(time.time_ns()))
    folder.mkdir(parents=True)
    binary = folder / f'{case}_{mode}_source.bin'
    evidence = export(source, case, mode, binary)
    _, rec = driver.load_model(case)
    if inputs['source_case'] != rec['source_case']:
        raise ValueError('Native input differs from original Python physical case')
    workbook = folder / 'PythonOriginalField_Test.xlsm'
    shutil.copyfile(ROOT.parent/'G1R17/delivery/RPFEM_20261012_G1R17_P06N_Coarse_Checked.xlsm',workbook)
    shutil.copyfile(ROOT.parent/'G1R17/delivery/rpfem_native_x64.dll',folder/'rpfem_native_x64.dll')
    prior = {p.pid for p in psutil.process_iter() if p.name().upper() == 'EXCEL.EXE'}
    app = win32com.client.DispatchEx('Excel.Application')
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    assert pid not in prior, 'Refusing existing Excel process'
    record = dict(status='RUNNING',pid=pid,prior_pids=sorted(prior),case=case,mode=mode,
                  output=str(folder),source=evidence,dialogs=[],checks={},
                  input_sha256=hashlib.sha256(inputs_path.read_bytes()).hexdigest(),input_path=str(inputs_path.relative_to(REPO)),
                  harness_sha256=hashlib.sha256(Path(__file__).read_bytes()).hexdigest(),
                  dll_sha256=hashlib.sha256((folder/'rpfem_native_x64.dll').read_bytes()).hexdigest())
    book = None
    stop = threading.Event()
    def save():
        (folder/'summary.json').write_text(json.dumps(record,ensure_ascii=False,indent=2)+'\n',encoding='utf-8')
    def monitor():
        while not stop.wait(.2):
            def window(hwnd,_):
                if win32process.GetWindowThreadProcessId(hwnd)[1] != pid or not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd) != '#32770':
                    return
                texts, buttons = [], []
                def child(item,_):
                    text = win32gui.GetWindowText(item)
                    if text: texts.append(text)
                    if text in ('OK','&OK') or (win32gui.GetWindowText(hwnd)=='Microsoft Visual Basic' and ('(&E)' in text or text=='&End')):
                        buttons.append(item)
                win32gui.EnumChildWindows(hwnd,child,None)
                if buttons:
                    record['dialogs'].append(texts)
                    for button in buttons: win32gui.PostMessage(button,win32con.BM_CLICK,0,0)
            win32gui.EnumWindows(window,None)
    threading.Thread(target=monitor,daemon=True).start()
    try:
        app.Visible=False; app.DisplayAlerts=False; app.EnableEvents=False
        record['excel_version']=app.Version; record['automation_security']=app.AutomationSecurity
        book=app.Workbooks.Open(str(workbook),UpdateLinks=0)
        modules=list(sorted((ROOT/'src/vba').glob('*.bas')))+list(sorted((ROOT.parent/'G1R18_RayGate/src/vba').glob('*.bas')))
        modules += [Path(__file__).with_name('PythonWitnessProbe.bas')]
        record['source_sha256']={str(p.relative_to(REPO)):hashlib.sha256(p.read_bytes()).hexdigest() for p in modules}
        for path in modules:
            try: component=book.VBProject.VBComponents(path.stem)
            except Exception:
                component=book.VBProject.VBComponents.Add(1); component.Name=path.stem
            component.CodeModule.DeleteLines(1,component.CodeModule.CountOfLines)
            component.CodeModule.AddFromString(adapt.source_text(path))
        component=book.VBProject.VBComponents.Add(1);component.Name='AccuracyExport'
        component.CodeModule.AddFromString(adapt.source_text(REPO/'validation/G1R12_native25_20261009/scripts/tests/BatchEvidence.bas').replace('Private Sub BatchState(','Public Sub BatchState('))
        book.Activate(); book.VBProject.VBComponents('PythonWitnessProbe').CodeModule.CodePane.Show()
        control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled: control.Execute(); time.sleep(.5)
        if control.Enabled:
            pane=app.VBE.ActiveCodePane; line=pane.GetSelection()[0]
            record['compile_error']=dict(module=pane.CodeModule.Name,line=line,source=pane.CodeModule.Lines(max(1,line-2),5))
            raise RuntimeError('VBA compile failed: '+repr(record['compile_error']))
        app.VBE.MainWindow.Visible=False; record['compile']='PASS';save()
        call=lambda name,*values:app.Run("'"+book.Name+"'!"+name,*values)
        call('RPX_Setup')
        for sheet,ranges in [('要素定義',[('A','C','points',7),('E','H','regions',7),('J','T','faces',7)]),('材料データ',[('A','E','materials',5)])]:
            ws=book.Worksheets(sheet); ws.Range('A7:T1006' if sheet=='要素定義' else 'A5:Q1004').ClearContents()
            for col,last,key,first in ranges:
                rows=tuple(tuple(row) for row in inputs[key])
                ws.Range(f'{col}{first}:{last}{first+len(rows)-1}').Value2=rows
        settings={**inputs['settings'],'QUADRATURE':8,'TEST_INJECTION':0,'HSD_CAPTURE':0,'DIAGNOSTIC_MODE':1}
        ws=book.Worksheets('設定')
        setting_rows=ws.Range('B4:D999').Value2
        slots={str(row[2]):i+4 for i,row in enumerate(setting_rows) if row[2]}
        empty=[i+4 for i,row in enumerate(setting_rows) if not row[0] and not row[2]]
        for key,value in settings.items():
            row=slots.get(key)
            if row is None:
                row=empty.pop(0);slots[key]=row
                ws.Cells(row,4).Value2=key
            ws.Cells(row,2).Value2=value
        book.Worksheets('浸透入力').Range('D4').Value2=0
        record['settings']=settings
        raw=binary.read_bytes()
        nn,ne,nm=struct.unpack_from('<3i',raw,8)
        for label,offset,value,reason in [('q4',24,struct.pack('<i',4),'SCOPE_OR_DIMENSIONS'),
              ('nan_xy',68,struct.pack('<d',float('nan')),'NONFINITE_SOURCE'),
              ('changed_material',68+16*nn+20*ne,struct.pack('<d',1e-15),'MATERIAL_C')]:
            changed=bytearray(raw);changed[offset:offset+len(value)]=value
            path=folder/(label+'.bin');path.write_bytes(changed)
            result=str(call('PythonRejectWitnessProbe',str(path)))
            assert result=='REJECT;'+reason,result
            record['checks'][label]=result
        save(); print('START_NATIVE_FIELD_AUDIT',case,mode,folder,pid,flush=True)
        started=time.perf_counter()
        result=str(call('PythonOriginalFieldProbe',str(binary),mode,str(folder/'audited')))
        record['native_result']=result;record['audit_seconds']=time.perf_counter()-started
        save(); print(result,flush=True)
        imported=read_state((folder/'audited_native.bin').read_bytes())
        original=read_state(raw)
        for index in range(4): assert np.array_equal(imported[index],original[index]),'Native import changed geometry/material/body'
        for edge,item in original[4].items():
            got=imported[4][edge]
            assert got['kind']==item['kind'] and got['rigid']==item['rigid'] and np.array_equal(got['loads'],item['loads'])
        count=12*ne if mode=='upper' else len(original[5])
        assert np.array_equal(imported[5][:count],original[5][:count]),'Native import changed physical field'
        native_mesh=driver.load_saved_mesh(source/f'best_{mode}_mesh.npz',rec)
        audit=core.evaluate(core.Prepared(native_mesh,evidence['fs'],mode,8,reduction='none'),imported[5])
        record['independent_replay']=audit
        record['independent_vba_matrix_replay']=replay_original(folder,mode)
        record['geometry_material_load_field_roundtrip']='PASS'
        assert result.startswith('PASS;'),result
        assert audit['witness'] and audit['physical_pass'],audit
        record['status']='PASS';book.Save()
    except Exception as error:
        record['status']='FAIL';record['error']=str(error)
        raise
    finally:
        stop.set();save()
        if book is not None:book.Close(False)
        app.Quit();book=None;app=None;gc.collect()
    print(json.dumps({k:record[k] for k in ('status','pid','output','native_result')},ensure_ascii=False),flush=True)


if __name__=='__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--case',choices=['P03A','P03N','P06A','P06N'],required=True)
    parser.add_argument('--mode',choices=['lower','upper'],required=True)
    parser.add_argument('--folder',type=Path,required=True)
    args=parser.parse_args()
    run(args.case,args.mode,args.folder.resolve())
