"""Native Excel test helpers; never change Excel trust or macro security."""
from pathlib import Path
import os
import hashlib
import time

EXPECTED_SHA256 = '4584ab05212eda14fbd5a559d41781c66b55615a0e1a78605e0afd5025b86fd7'

def source_book():
    root = Path(__file__).resolve().parents[1]
    default = root / 'baseline' / 'RPFEM_20260930_G1R7_Robust_C5_Verified.xlsm'
    if not default.exists():
        default = Path.home() / 'Downloads' / 'RPFEM_20260930_G1R7_Robust_C5_Verified.xlsm'
    path = Path(os.environ.get('RPFEM_INPUT_XLSM', str(default))).resolve()
    if hashlib.sha256(path.read_bytes()).hexdigest() != EXPECTED_SHA256:
        raise ValueError('Input workbook differs from the protected G1R7 C5 baseline')
    return path

def compile_book(app, book):
    book.VBProject.VBComponents('RPX_Fs').Activate()
    ctl = app.VBE.CommandBars.FindControl(1, 578)
    if ctl.Enabled:
        ctl.Execute()
    if ctl.Enabled:
        import win32gui, win32process, win32con
        pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        def inspect(hwnd, _):
            if win32process.GetWindowThreadProcessId(hwnd)[1] != pid:
                return
            def child(c, _):
                value = win32gui.GetWindowText(c)
                if value:
                    print('DIALOG ' + value, flush=True)
                if value == 'OK':
                    win32gui.PostMessage(c, win32con.BM_CLICK, 0, 0)
            win32gui.EnumChildWindows(hwnd, child, None)
        win32gui.EnumWindows(inspect, None)
        time.sleep(0.5)
        pane = app.VBE.ActiveCodePane
        print('PANE ' + pane.CodeModule.Name + ' ' + str(pane.GetSelection()), flush=True)
        start=pane.GetSelection()[0]
        print('AT_ERROR '+repr(pane.CodeModule.Lines(max(1,start-3),7)),flush=True)
        (Path(__file__).resolve().parents[1]/'results/compile_failed_source.txt').write_text(pane.CodeModule.Lines(1,pane.CodeModule.CountOfLines),encoding='utf-8')
    assert not ctl.Enabled, 'VBA compile failed'
    print('COMPILE PASS', flush=True)
