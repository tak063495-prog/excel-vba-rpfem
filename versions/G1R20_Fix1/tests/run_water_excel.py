"""Owned Excel water lifecycle micro; full frozen63 modules, solver0."""
from pathlib import Path
import gc, hashlib, json, shutil, threading, time, traceback
import psutil, win32com.client, win32con, win32gui, win32process

ROOT = Path(__file__).resolve().parents[1]
REPO = ROOT.parents[1]
BASE = REPO / 'versions/G1R20/delivery'
OUT = ROOT / 'results' / ('water_cache_' + str(time.time_ns()))
OUT.mkdir(parents=True)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def code(path):
    return '\r\n'.join(v for v in path.read_text(encoding='utf8').splitlines()
                        if not v.startswith('Attribute VB_'))


def main():
    prior = [p.pid for p in psutil.process_iter() if p.name().upper() == 'EXCEL.EXE']
    if prior:
        raise RuntimeError('Another Excel exists: ' + repr(prior))
    modules = {p.stem:p for p in (ROOT / 'src/vba').glob('*.bas')}
    if len(modules) != 63:
        raise ValueError('Expected frozen full63 modules')
    probe = ROOT / 'tests/WaterCacheFix1Probe.bas'
    bridges = {'RPX_Water': ROOT / 'tests/WaterCacheFix1WaterBridge.bas.txt',
               'RPX_SeepUI': ROOT / 'tests/WaterCacheFix1Bridge.bas.txt'}
    paths = list(modules.values()) + list((ROOT / 'import').glob('*.bas')) + [probe] + list(bridges.values())
    paths += [BASE / 'RPFEM_20261016_G1R20.xlsm', BASE / 'rpfem_native_x64.dll', Path(__file__)]
    before = {str(p.resolve()):sha(p) for p in paths}
    bookpath = OUT / 'WaterCacheFix1Probe.xlsm'
    shutil.copyfile(BASE / 'RPFEM_20261016_G1R20.xlsm', bookpath)
    shutil.copyfile(BASE / 'rpfem_native_x64.dll', OUT / 'rpfem_native_x64.dll')
    shutil.copyfile(__file__, OUT / 'harness_snapshot.py')
    snapshots = OUT / 'loaded_source_snapshot'
    snapshots.mkdir()
    for name, path in modules.items():
        text = code(path)
        if name in bridges:
            text += '\r\n' + code(bridges[name])
        (snapshots / (name + '.bas')).write_bytes(text.encode('utf8'))
    shutil.copyfile(probe, snapshots / probe.name)
    rec = dict(status='PREPARING', solver_calls=0, numerical_certificate_created=False,
        source_sha256_before=before, prior_excel_pids=prior, product_modules=63,
        test_only_bridge_modules=list(bridges), dialogs=[],
        python_owner=dict(pid=psutil.Process().pid, create_time=psutil.Process().create_time(),
                          cmdline=psutil.Process().cmdline()), started_epoch=time.time())
    def save():
        (OUT / 'summary.json').write_text(json.dumps(rec, indent=2, ensure_ascii=False) + '\n', encoding='utf8')
    save()
    app = book = component = control = None
    stop = threading.Event()
    try:
        app = win32com.client.DispatchEx('Excel.Application')
        pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        owner = psutil.Process(pid)
        birth, command = owner.create_time(), owner.cmdline()
        rec['excel_owner'] = dict(pid=pid, create_time=birth, cmdline=command, exe=owner.exe())
        save()
        def watch():
            deadline = time.monotonic() + 170
            while not stop.wait(.1):
                if time.monotonic() > deadline:
                    if psutil.pid_exists(pid):
                        process = psutil.Process(pid)
                        if process.create_time() == birth and process.cmdline() == command:
                            process.kill()
                            rec['watchdog'] = 'EXACT_OWNED_170S_STOP'
                            save()
                    return
                def visit(window, unused):
                    if win32process.GetWindowThreadProcessId(window)[1] != pid or win32gui.GetClassName(window) != '#32770':
                        return
                    children = []
                    win32gui.EnumChildWindows(window, lambda c, _:children.append((c, win32gui.GetWindowText(c))), None)
                    words = [v for _, v in children]
                    if words not in rec['dialogs']:
                        rec['dialogs'].append(words)
                        save()
                    for child, label in children:
                        if label in ('OK', '&OK') or '&End' in label or '終了' in label:
                            win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
                win32gui.EnumWindows(visit, None)
        watcher = threading.Thread(target=watch, daemon=True)
        watcher.start()
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.MultiThreadedCalculation.Enabled = False
        book = app.Workbooks.Open(str(bookpath), UpdateLinks=0)
        for name in [c.Name for c in book.VBProject.VBComponents if c.Type == 1]:
            book.VBProject.VBComponents.Remove(book.VBProject.VBComponents(name))
        for path in sorted(snapshots.glob('*.bas')):
            component = book.VBProject.VBComponents.Add(1)
            component.Name = path.stem
            component.CodeModule.AddFromString(code(path))
        component.CodeModule.CodePane.Show()
        control = app.VBE.CommandBars.FindControl(1, 578)
        if control.Enabled:
            control.Execute()
            time.sleep(.2)
        if control.Enabled:
            pane = app.VBE.ActiveCodePane
            line = int(pane.GetSelection()[0])
            raise RuntimeError('VBA compile failed ' + pane.CodeModule.Name + ':' + pane.CodeModule.Lines(max(1, line-2), 5))
        rec['compile'] = 'PASS'
        rec['loaded_code_sha256'] = {c.Name:hashlib.sha256(c.CodeModule.Lines(1, c.CodeModule.CountOfLines).encode('utf8')).hexdigest()
                                    for c in book.VBProject.VBComponents if c.Type == 1}
        save()
        app.VBE.MainWindow.Visible = False
        rec['result'] = app.Run("'" + book.Name + "'!WaterCacheFix1Run", str(OUT / 'water.tsv'))
        rec['status'] = 'PASS' if str(rec['result']).startswith('PASS ') else 'FAIL'
        rec['water_tsv_sha256'] = sha(OUT / 'water.tsv')
        rec['test_rows'] = (OUT / 'water.tsv').read_text(encoding='cp932').splitlines()[1:]
        book.Save()
        save()
    except BaseException as error:
        rec.update(status='FAIL', error=repr(error), traceback=traceback.format_exc())
        save()
    finally:
        try:
            if book is not None:
                book.Close(False)
            if app is not None:
                app.Quit()
        except Exception as error:
            rec['close_error'] = repr(error)
        component = control = book = app = None
        gc.collect()
        stop.set()
        if 'watcher' in locals():
            watcher.join(2)
        if 'pid' in locals():
            for _ in range(50):
                if not psutil.pid_exists(pid) or psutil.Process(pid).create_time() != birth:
                    break
                time.sleep(.1)
            rec['owned_excel_exited'] = not psutil.pid_exists(pid) or psutil.Process(pid).create_time() != birth
        rec['source_sha256_after'] = {str(p.resolve()):sha(p) for p in paths}
        rec['source_unchanged'] = before == rec['source_sha256_after']
        rec['finished_epoch'] = time.time()
        save()
    print(json.dumps(dict(output=str(OUT), status=rec['status'], result=rec.get('result'),
        owned_excel_exited=rec.get('owned_excel_exited'), source_unchanged=rec['source_unchanged']), ensure_ascii=False), flush=True)
    return 0 if rec['status'] == 'PASS' and rec.get('owned_excel_exited') and rec['source_unchanged'] else 2


if __name__ == '__main__':
    raise SystemExit(main())
