from pathlib import Path
import time
import win32gui,win32process,win32con
ROOT=Path(__file__).resolve().parents[1]
def compile_book(app,book):
    book.Activate()
    book.VBProject.VBComponents('RPX_Native').Activate()
    book.VBProject.VBComponents('RPX_Native').CodeModule.CodePane.Show()
    control=app.VBE.CommandBars.FindControl(1,578)
    if control.Enabled: control.Execute(); time.sleep(.25)
    if control.Enabled:
        pid=win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        def inspect(hwnd,_):
            if win32process.GetWindowThreadProcessId(hwnd)[1]!=pid:return
            def child(c,_):
                text=win32gui.GetWindowText(c)
                if text.startswith('コンパイル'): print(text,flush=True)
                if text=='OK': win32gui.PostMessage(c,win32con.BM_CLICK,0,0)
            win32gui.EnumChildWindows(hwnd,child,None)
        win32gui.EnumWindows(inspect,None); time.sleep(.4)
        pane=app.VBE.ActiveCodePane; line=pane.GetSelection()[0]
        print('AT_ERROR',pane.CodeModule.Name,line,repr(pane.CodeModule.Lines(max(1,line-2),5)),flush=True)
        (ROOT/'results/compile_failed_source.txt').write_text(pane.CodeModule.Lines(1,pane.CodeModule.CountOfLines),encoding='utf-8')
    assert not control.Enabled,'VBA compile failed'
    app.VBE.MainWindow.Visible=False
    print('COMPILE PASS',flush=True)
