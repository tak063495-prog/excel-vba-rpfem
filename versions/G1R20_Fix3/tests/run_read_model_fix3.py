"""Isolated Excel regression for legacy model identity initialization."""
from pathlib import Path
import argparse
import gc
import hashlib
import json
import shutil
import time
import uuid

import psutil
import win32com.client
import win32process

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT.parent / "G1R20_Fix2"


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def source_code(path):
    return "\r\n".join(line for line in Path(path).read_text(encoding="utf8").splitlines()
                        if not line.startswith("Attribute VB_"))


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--phase", choices=["before", "fixed"], required=True)
    args = parser.parse_args()
    prior = [p.pid for p in psutil.process_iter() if p.name().upper() == "EXCEL.EXE"]
    if prior:
        raise RuntimeError("Another Excel exists: " + repr(prior))
    out = ROOT / "results" / ("read_model_" + args.phase + "_" + str(time.time_ns()))
    out.mkdir(parents=True)
    book_path = out / ("ReadModel_" + args.phase + ".xlsm")
    shutil.copyfile(BASE / "delivery/RPFEM_20261017_G1R20_Fix2.xlsm", book_path)
    shutil.copyfile(BASE / "delivery/rpfem_native_x64.dll", out / "rpfem_native_x64.dll")
    source_root = BASE if args.phase == "before" else ROOT
    rec = dict(state="RUNNING", phase=args.phase, token=uuid.uuid4().hex, prior_excel_pids=prior,
               candidate_only=True, fs_bound_attempts=0, completed_fs_runs=0, harness_sha256=sha(__file__),
               baseline_book_sha256=sha(BASE / "delivery/RPFEM_20261017_G1R20_Fix2.xlsm"),
               baseline_model_sha256=sha(BASE / "src/vba/RPX_Model.bas"),
               dll_sha256=sha(out / "rpfem_native_x64.dll"), source_modules={}, workbook_modules={}, calls={})
    for path in [Path(__file__), ROOT / "tests/ReadModelFix3Probe.bas",
                 ROOT / "tests/read_model_shear_c10.txt", ROOT / "tests/read_model_shear_c11.txt"]:
        shutil.copyfile(path, out / path.name)
    def save():
        (out / "summary.json").write_text(json.dumps(rec, indent=2, ensure_ascii=False,
                                                   allow_nan=False) + "\n", encoding="utf8")
    save()
    app = win32com.client.DispatchEx("Excel.Application")
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    rec.update(owned_excel_pid=pid, owned_excel_creation_time=psutil.Process(pid).create_time())
    save()
    book = component = control = None
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.MultiThreadedCalculation.Enabled = False
        book = app.Workbooks.Open(str(book_path), UpdateLinks=0)
        for path in sorted((source_root / "src/vba").glob("*.bas")):
            component = book.VBProject.VBComponents(path.stem)
            module = component.CodeModule
            if module.CountOfLines:
                module.DeleteLines(1, module.CountOfLines)
            module.AddFromString(source_code(path))
            rec["source_modules"][path.name] = sha(path)
            module = None
        if args.phase == "fixed":
            path=ROOT / "src/workbook/ThisWorkbook.code.txt"
            component=book.VBProject.VBComponents("ThisWorkbook")
            module=component.CodeModule
            if module.CountOfLines:
                module.DeleteLines(1,module.CountOfLines)
            module.AddFromString(source_code(path))
            rec["workbook_modules"]["ThisWorkbook"] = sha(path)
            shutil.copyfile(path,out / "ThisWorkbook.code.txt")
            shutil.copyfile(ROOT / "SOURCE_MANIFEST.json",out / "SOURCE_MANIFEST.json")
            module=None
        component = book.VBProject.VBComponents.Add(1)
        component.Name = "ReadModelFix3Probe"
        component.CodeModule.AddFromString(source_code(ROOT / "tests/ReadModelFix3Probe.bas"))
        book.Activate()
        component.CodeModule.CodePane.Show()
        control = app.VBE.CommandBars.FindControl(1, 578)
        if control.Enabled:
            control.Execute()
            time.sleep(.3)
        if control.Enabled:
            raise RuntimeError("Regression workbook compile failed")
        app.VBE.MainWindow.Visible = False
        rec["compile"] = "PASS"
        rec["settings"]={name:app.Run("'"+book.Name+"'!RPX_Setting",name,-1)
                         for name in ['DIRECT_DLL','ACCURACY_DLL_RECOVERY','DIRECT_DLL_SECONDS','QUADRATURE']}
        save()
        def call(label, function, *values):
            result = app.Run("'" + book.Name + "'!" + function, *values)
            rec["calls"][label] = result
            save()
            print(json.dumps(dict(label=label, result=result[:5]), ensure_ascii=False), flush=True)
            return result
        a = str(out / "read_model_shear_c10.txt")
        b = str(out / "read_model_shear_c11.txt")
        if args.phase == "before":
            solve = call("fresh_testfs", "RM3Solve", a, str(out / "fresh_fs.csv"), True)
            fresh = call("fresh_identity", "RM3Read", a, "fresh")
            stale = call("stale_identity", "RM3Read", a, "stale")
            if not solve[0].startswith("FAIL 91 "):
                raise AssertionError("Expected fresh TestFs runtime 91")
            if fresh[0] != "REJECTED" or fresh[1] != 91 or not fresh[3]:
                raise AssertionError("Expected fresh ExactModel runtime 91")
            if stale[0] != "RETURNED" or stale[4] != 1 or "9998:9999" not in stale[5]:
                raise AssertionError("Expected stale constraints in model identity")
            rec["expected_negative_reproduced"] = True
            rec["fs_bound_attempts"] = solve[3]
        else:
            reads = [call("fresh_identity", "RM3Read", a, "fresh"),
                     call("same_identity", "RM3Read", a, "keep"),
                     call("other_identity", "RM3Read", b, "keep"),
                     call("return_identity", "RM3Read", a, "keep"),
                     call("stale_identity", "RM3Read", a, "stale")]
            if any(row[0] != "RETURNED" or row[3] or row[4] != 0 for row in reads):
                raise AssertionError("Legacy model constraints must be initialized empty")
            if not reads[0][5] == reads[1][5] == reads[3][5] == reads[4][5]:
                raise AssertionError("Model identity retained previous model state")
            if reads[0][5] == reads[2][5]:
                raise AssertionError("Different material must have a different model identity")
            for label, model, expected, fresh in [("fresh_testfs", a, 10., True),
                                                   ("other_testfs", b, 11., False),
                                                   ("return_testfs", a, 10., False)]:
                row = call(label, "RM3Solve", model, str(out / (label + ".csv")), fresh)
                rec["fs_bound_attempts"] += row[3]
                if not row[0].startswith("PASS ") or abs(row[1] - expected) > .0001 or abs(row[2] - expected) > .0001:
                    raise AssertionError("Unexpected analytical Fs pair: " + repr(row))
                if row[4] != 8 or row[5] > 1e-7 or not row[6]:
                    raise AssertionError("Original q8 physical audit failed")
                if any(value > 1e-7 for value in row[7:11]) or row[11]:
                    raise AssertionError("Original raw physical audit or root completion failed")
                if row[12] != 'OPTIMAL' or row[16]:
                    raise AssertionError("Analytical Fs requires the exact optimum, not a direct recovery witness")
                rec["completed_fs_runs"] += 1
            if rec["settings"]['DIRECT_DLL'] != 1 or rec["settings"]['ACCURACY_DLL_RECOVERY'] != 1:
                raise AssertionError("Regression must keep the saved direct DLL profile enabled")
        rec["native_memory_maps"] = [m.path for m in psutil.Process(pid).memory_maps()
                                     if Path(m.path).name.lower().startswith("rpfem_")]
        expected_maps = [str(out / "rpfem_native_x64.dll")]
        if args.phase == "before" and not rec["native_memory_maps"]:
            rec["native_not_reached_before_identity_failure"] = True
        elif rec["native_memory_maps"] != expected_maps:
            raise AssertionError("Unexpected native DLL mapping")
        rec["state"] = "PASS"
    except BaseException as exc:
        rec.update(state="FAIL", error=repr(exc))
        save()
        raise
    finally:
        if psutil.pid_exists(pid):
            rec["native_memory_maps"]=[m.path for m in psutil.Process(pid).memory_maps()
                                       if Path(m.path).name.lower().startswith("rpfem_")]
        component = control = None
        if book is not None:
            book.Close(False)
        app.Quit()
        book = app = None
        gc.collect()
        for _ in range(100):
            if not psutil.pid_exists(pid):
                break
            time.sleep(.1)
        rec["owned_excel_exited"] = not psutil.pid_exists(pid)
        rec["baseline_book_unchanged"] = rec["baseline_book_sha256"] == sha(BASE / "delivery/RPFEM_20261017_G1R20_Fix2.xlsm")
        rec["baseline_model_unchanged"] = rec["baseline_model_sha256"] == sha(BASE / "src/vba/RPX_Model.bas")
        save()
    print(json.dumps(dict(state=rec["state"], output=str(out), owned_excel_exited=rec["owned_excel_exited"])), flush=True)


if __name__ == "__main__":
    main()
