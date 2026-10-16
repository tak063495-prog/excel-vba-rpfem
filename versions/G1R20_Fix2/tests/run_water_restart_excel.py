"""Owned Excel negative/positive save-close-process-restart water regression.

No numerical optimizer is used. Negative uses the released Fix1 product;
positive imports the frozen full63 Fix2 sources into its disposable copy.
"""
from pathlib import Path
import argparse
import gc
import hashlib
import json
import shutil
import threading
import time
import traceback
import xml.etree.ElementTree as ET
import zipfile

import psutil
import win32com.client
import win32con
import win32gui
import win32process

ROOT = Path(__file__).resolve().parents[1]
OLD = ROOT.parent / "G1R20_Fix1"
BOOK = OLD / "delivery/RPFEM_20261017_G1R20_Fix1.xlsm"
DLL = OLD / "delivery/rpfem_native_x64.dll"
PROBE = ROOT / "tests/WaterRestartFix2Probe.bas"
BRIDGE = OLD / "tests/WaterCacheFix1Bridge.bas.txt"


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def code(path):
    return "\r\n".join(line for line in Path(path).read_text(encoding="utf8").splitlines()
                         if not line.startswith("Attribute VB_"))


def property_snapshot(path, output):
    with zipfile.ZipFile(path) as archive:
        raw = archive.read("docProps/custom.xml") if "docProps/custom.xml" in archive.namelist() else b""
    output.write_bytes(raw)
    found = []
    if raw:
        for item in ET.fromstring(raw):
            if item.attrib.get("name") == "RPX_SeepResultWriteState":
                found.append(dict(type=list(item)[0].tag.rsplit("}", 1)[-1], value=list(item)[0].text or ""))
    return dict(sha256=sha(output), journal=found)


class Session:
    def __init__(self, rec, save, directory):
        self.rec, self.save, self.directory = rec, save, directory
        self.app = self.book = self.component = self.control = None
        self.stop = threading.Event()

    def __enter__(self):
        self.app = win32com.client.DispatchEx("Excel.Application")
        pid = win32process.GetWindowThreadProcessId(self.app.Hwnd)[1]
        owner = psutil.Process(pid)
        self.pid, self.birth, self.command = pid, owner.create_time(), owner.cmdline()
        self.rec["excel_owner"] = dict(pid=pid, create_time=self.birth, cmdline=self.command, exe=owner.exe())
        self.save()
        def watch():
            deadline = time.monotonic() + 120
            while not self.stop.wait(.1):
                if time.monotonic() > deadline:
                    try:
                        process = psutil.Process(pid)
                        if process.create_time() == self.birth and process.cmdline() == self.command:
                            process.kill()
                            self.rec["watchdog"] = "EXACT_OWNED_120S_STOP"
                            self.save()
                    except psutil.NoSuchProcess:
                        pass
                    return
                def visit(window, _):
                    if win32process.GetWindowThreadProcessId(window)[1] != pid or win32gui.GetClassName(window) != "#32770":
                        return
                    children = []
                    win32gui.EnumChildWindows(window, lambda child, _: children.append((child, win32gui.GetWindowText(child))), None)
                    words = [value for _, value in children]
                    if words not in self.rec["dialogs"]:
                        self.rec["dialogs"].append(words)
                        self.save()
                    for child, label in children:
                        if label in ("OK", "&OK") or "&End" in label or "終了" in label:
                            win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
                if "monitor_error" not in self.rec:
                    try:
                        win32gui.EnumWindows(visit, None)
                    except Exception as error:
                        if getattr(error, "winerror", None) == 1400 or (error.args and error.args[0] == 1400):
                            # An owned window can disappear between EnumWindows
                            # and its inspection while Excel closes normally.
                            self.rec["transient_window_1400_count"] = self.rec.get("transient_window_1400_count", 0) + 1
                        else:
                            self.rec["monitor_error"] = repr(error)
                        self.save()
        self.watcher = threading.Thread(target=watch, daemon=True)
        self.watcher.start()
        self.app.Visible = False
        self.app.DisplayAlerts = False
        self.app.EnableEvents = False
        self.app.MultiThreadedCalculation.Enabled = False
        return self

    def open(self, path):
        self.book = self.app.Workbooks.Open(str(path), UpdateLinks=0)

    def inject(self, positive):
        if positive:
            for name in [item.Name for item in self.book.VBProject.VBComponents if item.Type == 1]:
                self.book.VBProject.VBComponents.Remove(self.book.VBProject.VBComponents(name))
            for path in sorted((ROOT / "src/vba").glob("*.bas")):
                item = self.book.VBProject.VBComponents.Add(1)
                item.Name = path.stem
                item.CodeModule.AddFromString(code(path))
        item = self.book.VBProject.VBComponents("RPX_SeepUI")
        item.CodeModule.AddFromString(code(BRIDGE))
        self.component = self.book.VBProject.VBComponents.Add(1)
        self.component.Name = PROBE.stem
        self.component.CodeModule.AddFromString(code(PROBE))
        self.component.CodeModule.CodePane.Show()
        self.control = self.app.VBE.CommandBars.FindControl(1, 578)
        if self.control.Enabled:
            self.control.Execute()
            time.sleep(.2)
        if self.control.Enabled:
            pane = self.app.VBE.ActiveCodePane
            line = int(pane.GetSelection()[0])
            raise RuntimeError("VBA compile failed " + pane.CodeModule.Name + ":" + pane.CodeModule.Lines(max(1, line-2), 5))
        self.app.VBE.MainWindow.Visible = False
        self.rec["compile"] = "PASS"
        directory = self.directory / "loaded_source_snapshot"
        directory.mkdir()
        for item in self.book.VBProject.VBComponents:
            if item.Type == 1:
                (directory / (item.Name + ".bas")).write_bytes(item.CodeModule.Lines(1, item.CodeModule.CountOfLines).encode("utf8"))
        self.save()

    def call(self, name, *values):
        return list(self.app.Run("'" + self.book.Name + "'!" + name, *values))

    def __exit__(self, *_):
        try:
            if self.book is not None:
                self.book.Close(False)
            if self.app is not None:
                self.app.Quit()
        except Exception as error:
            self.rec["close_error"] = repr(error)
        self.component = self.control = self.book = self.app = None
        gc.collect()
        self.stop.set()
        self.watcher.join(2)
        for _ in range(50):
            if not psutil.pid_exists(self.pid) or psutil.Process(self.pid).create_time() != self.birth:
                break
            time.sleep(.1)
        self.rec["owned_excel_exited"] = not psutil.pid_exists(self.pid) or psutil.Process(self.pid).create_time() != self.birth
        self.save()


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--execute", action="store_true")
    parser.add_argument("--phase", choices=("all", "negative", "positive"), default="all")
    args = parser.parse_args()
    if not args.execute:
        parser.error("--execute is required")
    prior = [item.pid for item in psutil.process_iter() if item.name().upper() == "EXCEL.EXE"]
    if prior:
        raise RuntimeError("Another Excel exists: " + repr(prior))
    sources = list((ROOT / "src/vba").glob("*.bas"))
    assert len(sources) == 63
    paths = sources + list((ROOT / "import").glob("*.bas")) + [BOOK, DLL, PROBE, BRIDGE, Path(__file__)]
    before = {str(path.resolve()): sha(path) for path in paths}
    out = ROOT / "results" / ("water_restart_" + str(time.time_ns()))
    out.mkdir(parents=True)
    for path in [PROBE, BRIDGE, Path(__file__)]:
        shutil.copyfile(path, out / path.name)
    rec = dict(status="PREPARING", cases=[], solver_calls=0, dialogs=[], prior_excel_pids=prior,
               source_sha256_before=before, started_epoch=time.time(),
               python_owner=dict(pid=psutil.Process().pid, create_time=psutil.Process().create_time(), cmdline=psutil.Process().cmdline()))
    def save():
        (out / "summary.json").write_text(json.dumps(rec, ensure_ascii=False, indent=2) + "\n", encoding="utf8")
    save()
    try:
        phases = (False, True) if args.phase == "all" else (args.phase == "positive",)
        for positive in phases:
            for mode, failure in ((1, "writer"), (1, "ui"), (2, "writer"), (2, "ui")):
                directory = out / (f"{'positive' if positive else 'negative'}_mode{mode}_{failure}")
                directory.mkdir()
                path = directory / "WaterRestart.xlsm"
                shutil.copyfile(BOOK, path)
                shutil.copyfile(DLL, directory / DLL.name)
                row = dict(positive=positive, mode=mode, failure=failure, sessions=[])
                rec["cases"].append(row)
                for stage in ("prepare", "restart", "healthy_restart"):
                    session = dict(stage=stage, dialogs=[])
                    row["sessions"].append(session)
                    with Session(session, save, directory) as excel:
                        excel.open(path)
                        if stage == "prepare":
                            excel.inject(positive)
                            row["prepare"] = excel.call("WaterRestartPrepare", mode, failure)
                            assert row["prepare"][0] == "PASS", row
                            row["expected_pending_input"] = dict(head=row["prepare"][8], body_x=row["prepare"][9])
                        elif stage == "restart":
                            row["observe"] = excel.call("WaterRestartObserve")
                            if positive:
                                assert row["observe"][0] == "REJECTED" and row["observe"][1] == 5 and "SEEP_RESULT_WRITE_INCOMPLETE" in row["observe"][2], row
                                assert row["observe"][4] == "PENDING", row
                                assert row["observe"][5] == "REJECTED" and row["observe"][6] == 5, row
                            else:
                                assert row["observe"][0] == "ACCEPTED" and abs(row["observe"][1]-2.) < 1e-9 and row["observe"][2] == 0., row
                                assert row["observe"][4] == "ABSENT_LEGACY", row
                                if failure == "writer":
                                    assert abs(row["observe"][1] - row["expected_pending_input"]["head"]) > .1, row
                            row["recover"] = excel.call("WaterRestartRecover")
                            assert row["recover"][0] == "PASS", row
                            if positive:
                                assert row["recover"][6] == "COMMITTED|" + row["recover"][5], row
                                row["marker_checks"] = excel.call("WaterRestartMarkerChecks")
                                assert row["marker_checks"][0] == "PASS", row
                            row["healthy_before_restart"] = excel.call("WaterRestartObserve")
                            assert row["healthy_before_restart"][0] == "ACCEPTED", row
                        else:
                            row["healthy_after_restart"] = excel.call("WaterRestartObserve")
                            assert row["healthy_after_restart"][0] == "ACCEPTED", row
                            assert row["healthy_before_restart"] == row["healthy_after_restart"], row
                        excel.book.Save()
                    assert session["owned_excel_exited"] and not session["dialogs"] and "monitor_error" not in session, session
                    saved_path = directory / (stage + "_saved.xlsm")
                    shutil.copyfile(path, saved_path)
                    row[stage + "_saved"] = dict(sha256=sha(saved_path),
                        custom_properties=property_snapshot(saved_path, directory / (stage + "_custom.xml")))
                    if positive:
                        expected_marker = "PENDING" if stage == "prepare" else row["healthy_before_restart"][4]
                        assert row[stage + "_saved"]["custom_properties"]["journal"] == [dict(type="lpwstr", value=expected_marker)], row
                assert row["sessions"][0]["excel_owner"] != row["sessions"][1]["excel_owner"], "Process restart required"
                row["status"] = "PASS"
                save()
        if args.phase != "negative":
            # A healthy pre-Fix2 file has no property. Upgrade only product
            # modules after the actual old producer has saved its result.
            directory = out / "legacy_healthy_upgrade"
            directory.mkdir()
            path = directory / "LegacyHealthy.xlsm"
            shutil.copyfile(BOOK, path)
            shutil.copyfile(DLL, directory / DLL.name)
            row = dict(positive=True, legacy_upgrade=True, sessions=[])
            rec["cases"].append(row)
            for stage in ("old_producer", "new_reader"):
                session = dict(stage=stage, dialogs=[])
                row["sessions"].append(session)
                snapshot = directory / stage
                snapshot.mkdir()
                with Session(session, save, snapshot) as excel:
                    excel.open(path)
                    excel.inject(stage == "new_reader")
                    if stage == "old_producer":
                        row["prepare"] = excel.call("WaterRestartLegacy")
                        assert row["prepare"][0:2] == ["PASS", "ABSENT_LEGACY"], row
                    else:
                        row["observe"] = excel.call("WaterRestartObserve")
                        assert row["observe"][0] == "ACCEPTED" and row["observe"][1] == 2., row
                        assert row["observe"][4] == "ABSENT_LEGACY", row
                    excel.book.Save()
                assert session["owned_excel_exited"] and not session["dialogs"] and "monitor_error" not in session, session
            row["status"] = "PASS"
        rec["status"] = "PASS"
    except BaseException as error:
        rec.update(status="FAIL", error=repr(error), traceback=traceback.format_exc())
    rec["owned_excel_exited"] = all(session.get("owned_excel_exited", False) for row in rec["cases"] for session in row["sessions"])
    rec["source_sha256_after"] = {str(path.resolve()): sha(path) for path in paths}
    rec["source_unchanged"] = before == rec["source_sha256_after"]
    rec["finished_epoch"] = time.time()
    save()
    print(json.dumps(dict(output=str(out), status=rec["status"], cases=len(rec["cases"]), owned_excel_exited=rec["owned_excel_exited"], source_unchanged=rec["source_unchanged"]), ensure_ascii=False), flush=True)
    return 0 if rec["status"] == "PASS" and rec["owned_excel_exited"] and rec["source_unchanged"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
