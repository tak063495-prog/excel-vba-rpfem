"""Owned Excel integration of the full original-program accuracy DLL bridge."""
from pathlib import Path
import gc
import argparse
import hashlib
import importlib.util
import json
import shutil
import struct
import threading
import time

import numpy as np
from scipy import sparse
import psutil
import win32com.client
import win32con
import win32gui
import win32process

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT.parent / 'G1R17'


def source_text(path):
    return '\r\n'.join(line for line in path.read_text(encoding='utf-8').splitlines() if not line.startswith('Attribute VB_'))


def replay(path, dual=False):
    """Replay literal exported CSR with independent NumPy/scipy calculations."""
    import struct
    data = path.read_bytes()
    assert data[:8] == b'AC18P001'
    n, m, nnz, nc, neq = struct.unpack_from('<5i', data, 8)
    offset = 28
    def read(dtype, count):
        nonlocal offset
        result = np.frombuffer(data, dtype=dtype, count=count, offset=offset).copy()
        offset += result.nbytes
        return result
    ptr, ids, val, rhs, cost = read('<i4', m+1), read('<i4', nnz), read('<f8', nnz), read('<f8', m), read('<f8', n)
    cones = read('<i4', 2*nc).reshape(-1, 2)
    fs = read('<f8', 1).item(); q = read('<i4', 1).item()
    x = read('<f8', m if dual else n)
    assert offset == len(data) and q == 8 and np.isfinite(x).all()
    assert np.array_equal(np.diff(ptr) >= 0, np.ones(m, dtype=bool))
    A = sparse.csr_matrix((val, ids, ptr), shape=(m, n))
    if dual:
        stat = float(np.max(np.abs(cost + A.T @ x), initial=0.))
        slack = x.copy(); objective = -float(rhs @ x)
    else:
        slack = rhs - A @ x; objective = float(cost @ x)
        stat = float(np.max(np.abs(slack[:neq]), initial=0.))
    row=neq; cone=0.
    for kind, dim in cones[1:]:
        block=slack[row:row+dim]
        cone=max(cone, -block[0] if kind == 1 else float(np.hypot(block[1], block[2]) - block[0]))
        row+=dim
    assert row==m and stat <= 1e-7 and cone <= (1e-8 if dual else 1e-7)
    return dict(variables=n, rows=m, nonzeros=nnz, q=q, fs=fs, raw_equality_or_stationarity=stat, raw_cone=cone, original_objective=objective)


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--fs',action='store_true')
    args=parser.parse_args()
    folder = ROOT / 'results' / ('accuracy_native_excel_' + str(time.time_ns()))
    folder.mkdir(parents=True)
    path = folder / 'AccuracyNative_Test.xlsm'
    shutil.copyfile(BASE / 'delivery/RPFEM_20261012_G1R17_P06N_Coarse_Checked.xlsm', path)
    for source in [BASE / 'delivery/rpfem_native_x64.dll', ROOT / 'native/accuracy/bin/rpfem_accuracy_x64.dll']:
        shutil.copyfile(source, folder / source.name)
    prior = {p.pid for p in psutil.process_iter() if p.name().upper() == 'EXCEL.EXE'}
    app = win32com.client.DispatchEx('Excel.Application')
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    assert pid not in prior
    record = dict(status='RUNNING', pid=pid, prior_pids=sorted(prior), output=str(folder), dialogs=[])
    record['dll_sha256']={p.name:hashlib.sha256(p.read_bytes()).hexdigest() for p in folder.glob('*.dll')}
    record['harness_sha256']=hashlib.sha256(Path(__file__).read_bytes()).hexdigest()
    def save():
        (folder / 'summary.json').write_text(json.dumps(record, indent=2, ensure_ascii=False)+'\n', encoding='utf-8')
    stop=threading.Event()
    def monitor():
        while not stop.wait(.2):
            def visit(hwnd, _):
                if win32process.GetWindowThreadProcessId(hwnd)[1] != pid or not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd) != '#32770': return
                texts=[]
                win32gui.EnumChildWindows(hwnd, lambda child, _: texts.append(win32gui.GetWindowText(child)), None)
                if texts not in record['dialogs']: record['dialogs'].append(texts); save()
                def end_error(child, _):
                    if '終了' in win32gui.GetWindowText(child): win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
                win32gui.EnumChildWindows(hwnd,end_error,None)
                win32gui.PostMessage(hwnd, win32con.WM_KEYDOWN, win32con.VK_RETURN, 0)
            win32gui.EnumWindows(visit, None)
    watcher=threading.Thread(target=monitor, daemon=True); watcher.start(); save()
    book=None
    try:
        app.Visible=False; app.DisplayAlerts=False; app.EnableEvents=False
        book=app.Workbooks.Open(str(path), UpdateLinks=0)
        files=list(sorted((ROOT/'src/vba').glob('*.bas')))+[ROOT/'tests/AccuracyNativeProbe.bas']
        record['source_sha256']={str(p.relative_to(ROOT)):hashlib.sha256(p.read_bytes()).hexdigest() for p in files}
        for file in files:
            try: component=book.VBProject.VBComponents(file.stem)
            except Exception: component=book.VBProject.VBComponents.Add(1); component.Name=file.stem
            module=component.CodeModule; module.DeleteLines(1,module.CountOfLines); module.AddFromString(source_text(file))
        book.VBProject.VBComponents('RPX_AccuracyNative').CodeModule.AddFromString('''
Public Function RPX_AccuracyProbeEquality(ByRef original As RPX_AccuracyProgram, ByRef candidate() As Double, ByRef objective As Double) As Boolean
    Dim raw() As Double
    RPX_AccuracyProbeEquality = RecoverEquilibrium(original, candidate, objective, raw)
End Function
''')
        if args.fs:
            fs=book.VBProject.VBComponents('RPX_Fs').CodeModule
            fs.AddFromString('''
Public Function RPX_FsAccuracyProbeValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double) As Double
    RPX_FsAccuracyProbeValue = FsValue(factor, mode, target)
End Function
Public Function RPX_FsAccuracyProbeDetail(ByVal side As String) As String
    Dim stored As FsBracketState
    If side = "lower" Then stored = lowerBracket Else stored = upperBracket
    RPX_FsAccuracyProbeDetail = "status_a=" & stored.statusA & ";status_b=" & stored.statusB & ";guide_a=" & stored.originalGuideA & ";guide_b=" & stored.originalGuideB & ";recovered=" & stored.recoveredWitnessA
End Function
Public Function RPX_FsAccuracyProbeMetadata(ByVal side As String, ByVal endpoint As Double) As Boolean
    Dim saved As FsBracketState, a As Double, b As Double
    If side = "lower" Then saved = lowerBracket Else saved = upperBracket
    If side = "lower" Then lowerBracket.inputEpoch = saved.inputEpoch - 1 Else upperBracket.inputEpoch = saved.inputEpoch - 1
    If RPX_ActualRootBracket(side, a, b, endpoint) Then err.Raise 5, , "STALE_ACCURACY_ROOT_ACCEPTED"
    If side = "lower" Then lowerBracket = saved: lowerBracket.physicalKey = saved.physicalKey & "foreign" Else upperBracket = saved: upperBracket.physicalKey = saved.physicalKey & "foreign"
    If RPX_ActualRootBracket(side, a, b, endpoint) Then err.Raise 5, , "FOREIGN_ACCURACY_ROOT_ACCEPTED"
    If side = "lower" Then lowerBracket = saved: lowerBracket.b = saved.a + 0.01 Else upperBracket = saved: upperBracket.b = saved.a + 0.01
    If RPX_ActualRootBracket(side, a, b, endpoint) Then err.Raise 5, , "WIDE_ACCURACY_ROOT_ACCEPTED"
    If side = "lower" Then lowerBracket = saved: lowerBracket.target = saved.target + 0.000001 Else upperBracket = saved: upperBracket.target = saved.target + 0.000001
    If RPX_ActualRootBracket(side, a, b, endpoint) Then err.Raise 5, , "FOREIGN_ACCURACY_ROOT_TARGET_ACCEPTED"
    If side = "lower" Then lowerBracket = saved: lowerBracket.target = AccuracyNativeNonfinite(&H7FF80000) Else upperBracket = saved: upperBracket.target = AccuracyNativeNonfinite(&H7FF80000)
    If RPX_ActualRootBracket(side, a, b, endpoint) Then err.Raise 5, , "NAN_ACCURACY_ROOT_TARGET_ACCEPTED"
    If side = "lower" Then lowerBracket = saved: lowerBracket.mode = "upper" Else upperBracket = saved: upperBracket.mode = "lower"
    If RPX_ActualRootBracket(side, a, b, endpoint) Then err.Raise 5, , "FOREIGN_ACCURACY_ROOT_SIDE_ACCEPTED"
    If side = "lower" Then lowerBracket = saved Else upperBracket = saved
    RPX_FsAccuracyProbeMetadata = RPX_ActualRootBracket(side, a, b, endpoint)
End Function
''')
            component=book.VBProject.VBComponents.Add(1); component.Name='AccuracyNativeFsProbe'
            component.CodeModule.AddFromString(source_text(ROOT/'tests/AccuracyNativeFsProbe.bas'))
        book.Activate(); book.VBProject.VBComponents('RPX_AccuracyNative').CodeModule.CodePane.Show()
        control=app.VBE.CommandBars.FindControl(1,578)
        if control.Enabled: control.Execute(); time.sleep(.5)
        if control.Enabled:
            pane=app.VBE.ActiveCodePane; line=pane.GetSelection()[0]
            record['compile_error']=dict(module=pane.CodeModule.Name,line=line,source=pane.CodeModule.Lines(max(1,line-2),5))
            raise RuntimeError('VBA compile failed: '+repr(record['compile_error']))
        app.VBE.MainWindow.Visible=False; record['compile']='PASS'; save()
        print('G1R18 accuracy bridge real Excel compile PASS',flush=True)
        result=app.Run("'"+book.Name+"'!"+('AccuracyNativeFsUnit' if args.fs else 'AccuracyNativeUnit'),str(folder))
        record['result']=result
        assert str(result).startswith('PASS'), result
        record['independent_original_replay']={p.name:replay(p) for p in folder.glob('*_recovered.bin')}
        record['independent_original_dual_replay']={p.name:replay(p,dual=True) for p in folder.rglob('*.accuracy_dual_*.bin')}
        assert len(record['independent_original_dual_replay'])>0
        for name,item in record['independent_original_dual_replay'].items():
            assert item['original_objective'] >= (1.000003 if name.endswith('_upper.bin') else -.999998)
        old_proofs=list(folder.rglob('*.original_lower_dual_*.bin'))
        if old_proofs:
            spec=importlib.util.spec_from_file_location('guide_replay',ROOT.parent/'G1R18_RayGate/tests/replay_original_guide.py')
            module=importlib.util.module_from_spec(spec);spec.loader.exec_module(module)
            record['original_lp_dual_replay']={p.name:module.replay_dual(p,lower=True) for p in old_proofs}
        if args.fs:
            exact=(folder/'own_root_endpoints.bin').read_bytes()
            assert exact[:8]==b'AC18R001' and len(exact)==40
            a,b,upper_a,upper_b=struct.unpack('<4d',exact[8:])
            assert b-a <= 1e-4*max(1.,(a+b)/2)
            assert any(row['fs']==b for row in record['independent_original_dual_replay'].values()),'Missing same-Fs native original negative proof'
            assert record['independent_original_replay']['root_lower_recovered.bin']['fs']==a
            record['same_mesh_signed_own_root']=dict(a=a,b=b,width=b-a,positive='original physical field',negative='independent full original dual',q=8)
            assert upper_b-upper_a<=1e-4*max(1.,(upper_a+upper_b)/2)
            assert any(row['fs']==upper_a for row in record['independent_original_dual_replay'].values())
            assert record['independent_original_replay']['root_upper_recovered.bin']['fs']==upper_b
            record['same_mesh_signed_upper_root']=dict(a=upper_a,b=upper_b,width=upper_b-upper_a,positive='independent full original dual',negative='original physical field',q=8)
        book.Save(); record['status']='PASS'; save()
        print(json.dumps({k:v for k,v in record.items() if k not in ['source_sha256','dialogs']},ensure_ascii=False),flush=True)
    except BaseException as exc:
        record['status']='FAIL'; record['error']=str(exc); save(); raise
    finally:
        try:
            if book is not None: book.Close(SaveChanges=False)
            app.Quit()
        finally:
            stop.set(); watcher.join(timeout=2); book=None; app=None; component=None; module=None; control=None; fs=None; pane=None; gc.collect()
        for _ in range(50):
            if not psutil.pid_exists(pid): break
            time.sleep(.1)
        record['owned_excel_exited']=not psutil.pid_exists(pid); save()


if __name__=='__main__': main()
