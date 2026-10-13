"""Serial isolated real Excel checks without changing macro/VBA security policy."""
from pathlib import Path
import argparse
import gc
import json
import shutil
import threading
import time

import psutil
import win32com.client
import win32con
import win32gui
import win32process

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT.parent / "G1R17"


def source_text(path):
    return "\r\n".join(line for line in path.read_text(encoding="utf-8").splitlines() if not line.startswith("Attribute VB_"))


def read_log(path):
    data = path.read_bytes()
    return data.decode("utf-16" if data[:2] in (b"\xff\xfe", b"\xfe\xff") else "utf-8-sig")


def verify_logs(folder):
    logs = list((folder / "RPFEM_logs").glob("*.tsv"))
    continuation = [p for p in logs if "accuracy_begin" in read_log(p)]
    assert len(continuation) == 1, "Expected continuation only for the enabled probe"
    events = read_log(continuation[0])
    assert "accuracy_round" in events, "No completed accuracy refinement"
    assert "fs_tolerance=0.0001" in events, "Final Fs tolerance changed"
    return {"continuation_log": str(continuation[0]), "continuation_checks": "PASS;enabled_only;completed_round;final_tolerance_preserved"}


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--compile-only", action="store_true")
    parser.add_argument("--failure-only", action="store_true")
    parser.add_argument("--verify-folder", type=Path)
    args = parser.parse_args()
    if args.verify_folder:
        folder = args.verify_folder.resolve()
        evidence = json.loads((folder / "summary.json").read_text(encoding="utf-8"))
        assert all("physical_gates=PASS" in row["result"] for row in evidence["runs"].values())
        assert set(evidence["runs"]) == {"disabled", "continued"}
        verification = {"status": "PASS", "scope": "Reinspect retained real Excel runs; no rerun or source change", "previous_harness_error": evidence.get("error"), **verify_logs(folder)}
        (folder / "log_verification.json").write_text(json.dumps(verification, indent=2) + "\n", encoding="utf-8")
        print(json.dumps(verification), flush=True)
        return
    folder = ROOT / "results" / ("adapt_excel_" + str(time.time_ns()))
    folder.mkdir(parents=True)
    path = folder / "AdaptAccuracy_Test.xlsm"
    shutil.copyfile(BASE / "delivery/RPFEM_20261012_G1R17_P06N_Coarse_Checked.xlsm", path)
    shutil.copyfile(BASE / "delivery/rpfem_native_x64.dll", folder / "rpfem_native_x64.dll")
    prior = {p.pid for p in psutil.process_iter() if p.name().upper() == "EXCEL.EXE"}
    app = win32com.client.DispatchEx("Excel.Application")
    pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
    assert pid not in prior, "Refusing to use an existing Excel process"
    record = {"status": "RUNNING", "pid": pid, "prior_pids": sorted(prior), "output": str(folder), "runs": {}, "dialogs": []}
    book = None
    stop = threading.Event()

    def save():
        (folder / "summary.json").write_text(json.dumps(record, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")

    def monitor():
        while not stop.wait(0.2):
            def window(hwnd, _):
                if win32process.GetWindowThreadProcessId(hwnd)[1] != pid or not win32gui.IsWindowVisible(hwnd) or win32gui.GetClassName(hwnd) != "#32770":
                    return
                texts, buttons = [], []
                def child(item, _):
                    text = win32gui.GetWindowText(item)
                    if text:
                        texts.append(text)
                    if text in ("OK", "&OK"):
                        buttons.append(item)
                win32gui.EnumChildWindows(hwnd, child, None)
                if buttons:
                    record["dialogs"].append(texts)
                    for item in buttons:
                        win32gui.PostMessage(item, win32con.BM_CLICK, 0, 0)
            win32gui.EnumWindows(window, None)
    threading.Thread(target=monitor, daemon=True).start()
    try:
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        record["excel_version"] = app.Version
        record["automation_security"] = app.AutomationSecurity
        book = app.Workbooks.Open(str(path), UpdateLinks=0)
        for name in ("RPX_Adapt", "RPX_Seep", "RPX_SeepUI"):
            module = book.VBProject.VBComponents(name).CodeModule
            module.DeleteLines(1, module.CountOfLines)
            module.AddFromString(source_text(ROOT / "src/vba" / (name + ".bas")))
        probe = book.VBProject.VBComponents.Add(1)
        probe.Name = "AdaptAccuracyProbe"
        probe.CodeModule.AddFromString(source_text(ROOT / "tests/AdaptAccuracyProbe.bas"))
        book.Activate()
        book.VBProject.VBComponents("RPX_Adapt").CodeModule.CodePane.Show()
        control = app.VBE.CommandBars.FindControl(1, 578)
        if control.Enabled:
            control.Execute()
            time.sleep(0.5)
        if control.Enabled:
            pane = app.VBE.ActiveCodePane
            line = pane.GetSelection()[0]
            record["compile_error"] = {"module": pane.CodeModule.Name, "line": line, "source": pane.CodeModule.Lines(max(1, line - 2), 5)}
            raise RuntimeError("VBA compilation failed: " + repr(record["compile_error"]))
        app.VBE.MainWindow.Visible = False
        record["compile"] = "PASS"
        print("G1R18 adaptive and Issue8 candidate compile PASS", flush=True)
        save()
        call = lambda name, *values: app.Run("'" + book.Name + "'!" + name, *values)
        if not args.compile_only:
            cases = [("disabled", 64, 0, False), ("injected_failure", 256, 2, True)] if args.failure_only else [("disabled", 64, 0, False), ("continued", 256, 2, False)]
            for label, cap, rounds, inject_failure in cases:
                trace = folder / (label + ".tsv")
                print("START", label, flush=True)
                started = time.perf_counter()
                result = str(call("AccuracyProbeRun", cap, rounds, str(trace), inject_failure))
                assert result.startswith("PASS;"), result
                record["runs"][label] = {"result": result, "seconds": time.perf_counter() - started}
                print("END", label, result, flush=True)
                save()
            book.Save()
            if args.failure_only:
                def values(label):
                    return dict(item.split("=", 1) for item in record["runs"][label]["result"].split(";")[1:])
                assert values("disabled") == values("injected_failure"), "Failure did not retain the exact baseline bounds and witness meshes"
                logs = [read_log(p) for p in (folder / "RPFEM_logs").glob("*.tsv")]
                assert any("reason=numerical_or_geometry_failure" in log and "CONIC_NOT_CONVERGED" in log for log in logs)
                assert not any("\taccuracy_round\t" in log for log in logs), "Failed trial was counted as an accepted refinement"
                record["failure_checks"] = "PASS;exact_baseline_retained;both_original_field_gates;failed_trial_not_accepted"
            else:
                record.update(verify_logs(folder))
        book.Save()
        record["status"] = "PASS"
    except Exception as error:
        record["status"] = "FAIL"
        record["error"] = str(error)
        raise
    finally:
        stop.set()
        save()
        if book is not None:
            book.Close(False)
        app.Quit()
        book = None
        app = None
        gc.collect()
    print(json.dumps(record, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
