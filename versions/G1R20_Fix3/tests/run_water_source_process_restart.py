"""Two distinct owned Excel processes: public mode2 source survives restart."""
from pathlib import Path
import argparse
import gc
import importlib.util
import json
import shutil
import threading
import time
import traceback

import psutil
import win32com.client
import win32con
import win32gui
import win32process

ROOT = Path(__file__).resolve().parents[1]
SPEC = importlib.util.spec_from_file_location('source_util', ROOT.parent/'G1R20_Fix1/tests/UI_run_excel.py')
UTIL = importlib.util.module_from_spec(SPEC)
SPEC.loader.exec_module(UTIL)
sha, code, code_key = UTIL.sha, UTIL.code, UTIL.code_key


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--execute', action='store_true')
    parser.add_argument('--output', required=True, type=Path)
    args = parser.parse_args()
    if not args.execute:
        raise SystemExit('Requires coordinated --execute')
    out = args.output.resolve()
    out.mkdir(parents=True, exist_ok=False)
    template = ROOT/'delivery/RPFEM_20261017_G1R20_Fix3.xlsm'
    sources = sorted((ROOT/'src/vba').glob('*.bas'))
    document = ROOT/'src/workbook/ThisWorkbook.code.txt'
    probe = ROOT/'tests/WaterSourceFix3Probe.bas'
    monitored = [template, document, probe, Path(__file__), *sources,
                 *sorted((ROOT/'import').glob('*.bas')), ROOT/'import/ThisWorkbook.code.txt']
    before = {str(p): sha(p) for p in monitored}
    snapshot = out/'source_snapshot'
    snapshot.mkdir()
    for p in [*sources, document, probe]:
        shutil.copyfile(p, snapshot/p.name)
    shutil.copyfile(__file__, out/'runner_snapshot.py')
    path = out/'restart_mode2.xlsm'
    shutil.copyfile(template, path)
    rec = dict(status='PREPARING', stages=[], source_sha256_before=before,
               conic_optimizer_calls=0, public_mode2_producer_runs=0)

    def save():
        (out/'summary.json').write_text(json.dumps(rec, ensure_ascii=False, indent=2)+'\n', encoding='utf8')

    def run_stage(label):
        prior = [p.pid for p in psutil.process_iter() if p.name().upper() == 'EXCEL.EXE']
        if prior:
            raise RuntimeError('Excel slot occupied '+repr(prior))
        app = win32com.client.DispatchEx('Excel.Application')
        pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        process = psutil.Process(pid)
        birth, command = process.create_time(), process.cmdline()
        stage = dict(label=label, pid=pid, birth=birth, cmdline=command, prior_pids=prior,
                     dialogs=[], results={}, started=time.time())
        rec['stages'].append(stage)
        save()
        stop = threading.Event()
        started = time.monotonic()

        def monitor():
            while not stop.wait(.15):
                if time.monotonic()-started > 180:
                    try:
                        owned = psutil.Process(pid)
                        if owned.create_time() == birth and owned.cmdline() == command:
                            stage['owned_timeout'] = True
                            save()
                            owned.kill()
                    except psutil.NoSuchProcess:
                        pass
                    return

                def visit(window, _):
                    if win32process.GetWindowThreadProcessId(window)[1] != pid or win32gui.GetClassName(window) != '#32770':
                        return
                    children = []
                    win32gui.EnumChildWindows(window, lambda child, _: children.append((child, win32gui.GetWindowText(child))), None)
                    labels = [text for _, text in children]
                    if labels not in stage['dialogs']:
                        stage['dialogs'].append(labels)
                        save()
                    for child, text in children:
                        if text in ['OK', '&OK', '終了(&E)', '&End']:
                            win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
                try:
                    win32gui.EnumWindows(visit, None)
                except Exception as error:
                    if getattr(error, 'winerror', None) != 1400 and not (error.args and error.args[0] == 1400):
                        stage.setdefault('monitor_errors', []).append(repr(error))
                        save()

        watcher = threading.Thread(target=monitor, daemon=True)
        watcher.start()
        book = component = control = None
        try:
            app.Visible = False
            app.DisplayAlerts = False
            app.EnableEvents = False
            app.MultiThreadedCalculation.Enabled = False
            book = app.Workbooks.Open(str(path), UpdateLinks=0)
            for p in sources:
                component = book.VBProject.VBComponents(p.stem)
                assert code_key(component.CodeModule.Lines(1, component.CodeModule.CountOfLines)) == code_key(code(p)), p.name
            component = book.VBProject.VBComponents('ThisWorkbook')
            assert code_key(component.CodeModule.Lines(1, component.CodeModule.CountOfLines)) == code_key(code(document))
            if label == 'producer_process':
                component = book.VBProject.VBComponents.Add(1)
                component.Name = probe.stem
                component.CodeModule.AddFromString(code(probe))
            component = book.VBProject.VBComponents(probe.stem)
            assert code_key(component.CodeModule.Lines(1, component.CodeModule.CountOfLines)) == code_key(code(probe))
            component.CodeModule.CodePane.Show()
            control = app.VBE.CommandBars.FindControl(1, 578)
            if control.Enabled:
                control.Execute()
                time.sleep(.2)
            assert not control.Enabled, 'VBA compile failed'
            app.VBE.MainWindow.Visible = False
            call = lambda name, *values: list(app.Run("'"+book.Name+"'!"+name, *values))
            if label == 'producer_process':
                healthy = call('WaterSourcePublicRun', 2, False)
                assert healthy[0] == 'PASS', healthy
                rec['public_mode2_producer_runs'] += 1
                edit = call('WaterSourceEdit', 2)
                stage['results'] = dict(healthy=healthy, edit=edit,
                                        k2=book.Worksheets('浸透結果').Range('K2').Value2,
                                        k9=book.Worksheets('浸透結果').Range('K9').Value2)
                assert stage['results']['k2'] is True
            else:
                rejected = call('WaterSourceObserve')
                assert rejected[0] == 'REJECTED' and 'SEEP_SOURCE_CHANGED' in str(rejected[2]), rejected
                healthy = call('WaterSourcePublicRun', 2, False)
                assert healthy[0] == 'PASS', healthy
                rec['public_mode2_producer_runs'] += 1
                stage['results'] = dict(rejected_after_process_restart=rejected, recovered_public_producer=healthy)
            if stage['dialogs'] or stage.get('monitor_errors'):
                raise RuntimeError('Dialog or monitor error')
            book.Save()
            stage['status'] = 'PASS'
            save()
        finally:
            try:
                if book is not None:
                    book.Close(False)
            finally:
                try:
                    app.Quit()
                except Exception as error:
                    stage['quit_error'] = repr(error)
            book = component = control = app = process = None
            gc.collect()
            for _ in range(150):
                if not psutil.pid_exists(pid):
                    break
                time.sleep(.1)
            stop.set()
            watcher.join(2)
            stage['owned_excel_exited'] = not psutil.pid_exists(pid)
            stage['finished'] = time.time()
            save()
        assert stage['owned_excel_exited'], stage
        assert not stage.get('owned_timeout'), stage

    try:
        run_stage('producer_process')
        run_stage('consumer_new_process')
        assert rec['stages'][0]['pid'] != rec['stages'][1]['pid']
        rec['status'] = 'PASS'
    except BaseException as error:
        rec.update(status='FAIL', error=repr(error), traceback=traceback.format_exc())
        raise
    finally:
        rec['source_sha256_after'] = {str(p): sha(p) for p in monitored}
        rec['monitored_files_unchanged'] = rec['source_sha256_after'] == before
        rec['test_workbook_sha256'] = sha(path)
        if not rec['monitored_files_unchanged']:
            rec['status'] = 'FAIL'
        save()
    assert rec['status'] == 'PASS'
    print(json.dumps(dict(status=rec['status'], output=str(out),
                          owned_excel_pids=[s['pid'] for s in rec['stages']],
                          all_owned_processes_exited=all(s['owned_excel_exited'] for s in rec['stages'])), ensure_ascii=False))


if __name__ == '__main__':
    main()
