"""Serial isolated real Excel checks without changing macro/VBA security policy."""
from pathlib import Path
import argparse
import gc
import hashlib
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
    parser.add_argument("--fan-only", action="store_true")
    parser.add_argument("--fan-geometry-only", action="store_true")
    parser.add_argument("--quadrature-only", action="store_true")
    parser.add_argument("--ray-only", action="store_true")
    parser.add_argument("--ray-adapt-only", action="store_true")
    parser.add_argument("--fixed-recovery-only", action="store_true")
    parser.add_argument("--fixed-fs-only", action="store_true")
    parser.add_argument("--production-case", choices=("P06A", "P06N"))
    parser.add_argument("--target", type=int, default=1000)
    parser.add_argument("--cycles", type=int, default=8)
    parser.add_argument("--accuracy-cap", type=int, default=24000)
    parser.add_argument("--accuracy-rounds", type=int, default=8)
    parser.add_argument("--lower-target", type=int, default=4000)
    parser.add_argument("--fan-rounds", type=int, default=1)
    parser.add_argument("--verify-folder", type=Path)
    args = parser.parse_args()
    if args.fixed_fs_only:
        args.fixed_recovery_only = True
    if args.ray_adapt_only or args.fixed_recovery_only:
        args.ray_only = True
    ray_root = ROOT.parent / "G1R18_RayGate"
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
    record["dll_sha256"] = hashlib.sha256((folder / "rpfem_native_x64.dll").read_bytes()).hexdigest()
    record["harness_sha256"] = hashlib.sha256(Path(__file__).read_bytes()).hexdigest()
    if args.ray_only:
        record["probe_sha256"] = hashlib.sha256((ray_root / "tests/OriginalFarkasProbe.bas").read_bytes()).hexdigest()
        if args.fixed_recovery_only:
            record["fixed_probe_sha256"] = hashlib.sha256((ray_root / "tests/FixedLowerRecoveryProbe.bas").read_bytes()).hexdigest()
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
                    if text in ("OK", "&OK") or (win32gui.GetWindowText(hwnd) == "Microsoft Visual Basic" and ("(&E)" in text or text == "&End")):
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
        module_files = list(sorted((ROOT / "src/vba").glob("*.bas")))
        if args.ray_only:
            # The retained RayGate candidate predates the additive accuracy ABI.
            # Its own Conic type and Fs module must remain a coherent snapshot.
            module_files = [p for p in module_files if p.stem not in {"RPX_AccuracyNative", "RPX_Upper"}]
            module_files += [BASE / "src/vba/RPX_Upper.bas"]
            module_files += list(sorted((ray_root / "src/vba").glob("*.bas")))
        record["source_sha256"] = {str(p.relative_to(ROOT.parents[1])): hashlib.sha256(p.read_bytes()).hexdigest() for p in module_files}
        for file in module_files:
            name = file.stem
            try:
                component = book.VBProject.VBComponents(name)
            except Exception:
                component = book.VBProject.VBComponents.Add(1)
                component.Name = name
            module = component.CodeModule
            module.DeleteLines(1, module.CountOfLines)
            module.AddFromString(source_text(file))
        if args.ray_only:
            book.VBProject.VBComponents("RPX_Fs").CodeModule.AddFromString('''
Public Function RPX_FsProbeValue(ByVal factor As Double, ByVal mode As String, ByVal target As Double) As Double
    RPX_FsProbeValue = FsValue(factor, mode, target)
End Function
Public Function RPX_LowerFsProbeDetail() As String
    RPX_LowerFsProbeDetail = "status_a=" & lowerBracket.statusA & ";status_b=" & lowerBracket.statusB & ";guide_b=" & lowerBracket.originalGuideB & ";recovered_a=" & lowerBracket.recoveredWitnessA & ";a=" & lowerBracket.a & ";b=" & lowerBracket.b
End Function
Public Function RPX_LowerRootMetadataProbe(ByVal factor As Double) As Boolean
    Dim saved As FsBracketState, a As Double, b As Double, number As Long, message As String
    saved = lowerBracket
    On Error GoTo Failed
    lowerBracket.inputEpoch = saved.inputEpoch - 1
    If RPX_ActualRootBracket("lower", a, b, factor) Then err.Raise 5, , "FOREIGN_ROOT_EPOCH_ACCEPTED"
    lowerBracket = saved: lowerBracket.physicalKey = saved.physicalKey & "foreign"
    If RPX_ActualRootBracket("lower", a, b, factor) Then err.Raise 5, , "FOREIGN_ROOT_MODEL_ACCEPTED"
    lowerBracket = saved: lowerBracket.policy = saved.policy & "foreign"
    If RPX_ActualRootBracket("lower", a, b, factor) Then err.Raise 5, , "FOREIGN_ROOT_POLICY_ACCEPTED"
    lowerBracket = saved: lowerBracket.originalGuideB = False
    If RPX_ActualRootBracket("lower", a, b, factor) Then err.Raise 5, , "RECOVERED_ROOT_WITHOUT_FREE_DUAL_ACCEPTED"
    lowerBracket = saved: lowerBracket.b = saved.a + 0.002
    If RPX_ActualRootBracket("lower", a, b, factor) Then err.Raise 5, , "WIDE_RECOVERED_ROOT_ACCEPTED"
    lowerBracket = saved
    RPX_LowerRootMetadataProbe = RPX_ActualRootBracket("lower", a, b, factor)
    Exit Function
Failed:
    number = err.number: message = err.description: lowerBracket = saved
    err.Raise number, "RPX_LowerRootMetadataProbe", message
End Function
Public Function RPX_FsProbeDetail() As String
    RPX_FsProbeDetail = "valid=" & upperBracket.valid & ";mesh=" & upperBracket.meshKey & ";current_mesh=" & FsMeshKey() & ";q=" & upperBracket.subdivisions & ";current_q=" & RSubdivisions & ";status_a=" & upperBracket.statusA & ";status_b=" & upperBracket.statusB & ";guide_a=" & upperBracket.originalGuideA & ";a=" & upperBracket.a & ";b=" & upperBracket.b & ";fa=" & upperBracket.fa & ";fb=" & upperBracket.fb
End Function
''')
        probe = book.VBProject.VBComponents.Add(1)
        probe.Name = "AdaptAccuracyProbe"
        probe.CodeModule.AddFromString(source_text(ROOT / "tests/AdaptAccuracyProbe.bas"))
        export = book.VBProject.VBComponents.Add(1)
        export.Name = "AccuracyExport"
        export_text = source_text(ROOT.parents[1] / "validation/G1R12_native25_20261009/scripts/tests/BatchEvidence.bas")
        export.CodeModule.AddFromString(export_text.replace("Private Sub BatchState(", "Public Sub BatchState("))
        if args.ray_only:
            ray = book.VBProject.VBComponents.Add(1)
            ray.Name = "OriginalFarkasProbe"
            ray.CodeModule.AddFromString(source_text(ray_root / "tests/OriginalFarkasProbe.bas"))
            if args.fixed_recovery_only:
                fixed = book.VBProject.VBComponents.Add(1)
                fixed.Name = "FixedLowerRecoveryProbe"
                fixed.CodeModule.AddFromString(source_text(ray_root / "tests/FixedLowerRecoveryProbe.bas"))
                if args.fixed_fs_only:
                    fixed_fs = book.VBProject.VBComponents.Add(1)
                    fixed_fs.Name = "FixedLowerFsProbeModule"
                    fixed_fs.CodeModule.AddFromString(source_text(ray_root / "tests/FixedLowerFsProbe.bas"))
                    record["fixed_fs_probe_sha256"] = hashlib.sha256((ray_root / "tests/FixedLowerFsProbe.bas").read_bytes()).hexdigest()
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
        if args.ray_only:
            record["variant"] = "G1R18_RayGate"
            probes = [("OriginalFarkasUnitProbe", ()), ("C0WedgeGeometryProbe", ()), ("C0WedgeNativeProbe", (str(folder / "c0_wedge_native.tsv"),))]
            if args.ray_adapt_only:
                probes = [("C0WedgeAdaptProbe", (depth, inject, str(folder / label))) for depth, inject, label in [(0., False, "baseline.tsv"), (.01, False, "wedge.tsv"), (.01, True, "failure.tsv")]]
            if args.fixed_recovery_only:
                probes = [("FixedLowerRecoverySmallProbe", (str(folder / ("fixed_c0.tsv" if c0 else "fixed_cohesive.tsv")),c0)) for c0 in (False, True)]
            if args.fixed_fs_only:
                probes = [("FixedLowerFsProbe", (str(folder / "fixed_fs.tsv"),))]
            for name, values in probes:
                started = time.perf_counter()
                result = str(call(name, *values))
                assert result.startswith("PASS;"), result
                label = Path(values[0]).stem if args.fixed_recovery_only else (Path(values[-1]).stem if args.ray_adapt_only else name)
                record["runs"][label] = {"result": result, "seconds": time.perf_counter() - started}
                print(name, result, flush=True)
                save()
            native_logs = [read_log(p) for p in (folder / "RPFEM_logs").glob("*.tsv") if ".context." not in p.name]
            if args.fixed_recovery_only:
                assert any("fixed_lower_recovery" in t and "accepted=True" in t for t in native_logs)
                record["fixed_recovery_checks"] = "PASS;full_free_lambda_original_primal_and_field_gate;compiled_layout_cache_and_both_witnesses_retained;no_fixed_dual_or_root_capture"
                if args.fixed_fs_only:
                    record["fixed_recovery_checks"] = "PASS;positive_free_load_physical_witness;separate_original_free_load_negative_dual;fixed_dual_never_used;both_saved_bounds_fields_retained"
                    import importlib.util
                    spec = importlib.util.spec_from_file_location("lower_guide_replay", ray_root / "tests/replay_original_guide.py")
                    replay = importlib.util.module_from_spec(spec); spec.loader.exec_module(replay)
                    paths = sorted((folder / "RPFEM_logs").glob("*.original_lower_dual_*.bin"))
                    assert paths, "No original FREE-lambda dual proof retained"
                    duals = [replay.replay_dual(path, lower=True) for path in paths]
                    spec = importlib.util.spec_from_file_location("lower_primal_replay", ROOT / "tests/test_python_witness_excel.py")
                    primal = importlib.util.module_from_spec(spec); spec.loader.exec_module(primal)
                    checked = primal.replay_original(folder, "lower", "fixed_fs.tsv.free_original.bin", "fixed_fs.tsv.free_field.bin")
                    bounds = dict(part.split("=", 1) for part in result.split(";")[1:] if "=" in part)
                    a, b = float(bounds["root_a"]), float(bounds["root_b"])
                    assert a < b and b - a <= 1e-4 * max(1., (a + b) / 2)
                    assert any(abs(row["fs"] - b) < 1e-13 for row in duals), "Root B lacks its own original dual proof"
                    context = replay.context(folder / "fixed_fs.tsv.endpoint_context.tsv")
                    assert context["q"] == [["8"]] and abs(replay.exact_float(context["fs"][0][0]) - a) < 1e-13
                    record["independent_replay"] = dict(original_free_lower_duals=duals, original_physical_primal=checked, root_a=a, root_b=b, fixed_dual_used=False)
                    (folder / "original_lower_evidence_replay.json").write_text(json.dumps(record["independent_replay"], indent=2) + "\n", encoding="utf-8")
            elif args.ray_adapt_only:
                assert any("upper_c0_wedge_end" in t and "accepted=True" in t for t in native_logs)
                assert any("upper_c0_wedge_rejected" in t and "CONIC_NOT_CONVERGED" in t for t in native_logs)
                record["ray_checks"] = "PASS;wedge_own_q8_root_original_field_gate;lower_exact_retained;failed_wedge_old_witnesses_retained"
            else:
                assert len(native_logs) == 1, "Missing unique native original guide trace"
                assert "original_farkas_guide" in native_logs[0], "Native root did not exercise original Farkas acceptance"
                assert "ORIGINAL_FARKAS_GUIDE" in native_logs[0], "No committed original guide status"
                assert "root_width=" in native_logs[0], "Missing original endpoint bracket"
                import importlib.util
                specification = importlib.util.spec_from_file_location("replay_original_guide", ray_root / "tests/replay_original_guide.py")
                replay = importlib.util.module_from_spec(specification)
                specification.loader.exec_module(replay)
                record["independent_replay"] = replay.verify_folder(folder)
                record["ray_checks"] = "PASS;native_original_dual_gate;negative_original_field_gate;separate_q8_root;no_continuum_ray_capture"
        elif args.production_case:
            data = ROOT.parents[1] / "validation/accuracy_20261013/python/data" / (args.production_case + ".json")
            inputs = json.loads(data.read_text(encoding="utf-8"))
            record["case"] = args.production_case
            record["input_record_sha256"] = hashlib.sha256(data.read_bytes()).hexdigest()
            record["source_sha256"] = {p.name: hashlib.sha256(p.read_bytes()).hexdigest() for p in (ROOT / "src/vba").glob("*.bas")}
            record["dll_sha256"] = hashlib.sha256((folder / "rpfem_native_x64.dll").read_bytes()).hexdigest()
            call("RPX_Setup")
            for sheet, ranges in [("要素定義", [("A", "C", "points", 7), ("E", "H", "regions", 7), ("J", "T", "faces", 7)]), ("材料データ", [("A", "E", "materials", 5)])]:
                ws = book.Worksheets(sheet)
                ws.Range("A7:T1006" if sheet == "要素定義" else "A5:Q1004").ClearContents()
                for col, last, key, first in ranges:
                    rows = tuple(tuple(row) for row in inputs[key])
                    rng = ws.Range(f"{col}{first}:{last}{first + len(rows) - 1}")
                    rng.Value2 = rows
                    assert rng.Value2 == rows, key
            settings = {**inputs["settings"], "TARGET": args.target, "CYCLES": args.cycles,
                        "QUADRATURE": 8, "FS_TOL": .0001, "GAP": .01, "ADAPT": 1,
                        "ADAPT_MESH_POLICY": "ROBUST_REFINE", "ADAPT_ACCURACY_CAP": args.accuracy_cap,
                        "ADAPT_ACCURACY_ROUNDS": args.accuracy_rounds, "ADAPT_LOWER_TARGET": args.lower_target,
                        "ADAPT_LOWER_FAN_ROUNDS": args.fan_rounds, "ADAPT_CERT_QUADRATURE": 8,
                        "HSD_SOLVE_SECONDS": 300, "TEST_INJECTION": 0}
            ws = book.Worksheets("設定")
            for key, value in settings.items():
                slots = {str(ws.Cells(r, 4).Value2): r for r in range(4, 1000) if ws.Cells(r, 4).Value2}
                row = slots.get(key)
                if row is None:
                    row = next(r for r in range(4, 1000) if not ws.Cells(r, 4).Value2 and not ws.Cells(r, 2).Value2)
                    ws.Cells(row, 4).Value2 = key
                ws.Cells(row, 2).Value2 = value
            book.Worksheets("浸透入力").Range("D4").Value2 = 0
            record["settings"] = settings
            definition = str(call("AccuracyExport.BatchDefinition")).split("\t")
            assert list(map(int, definition[:3])) == [len(inputs[key]) for key in ("points", "regions", "faces")]
            assert definition[-1] == inputs["settings"]["ANALYSIS_POLICY"]
            assert abs(float(definition[3]) - inputs["source_case"]["area"]) < 1e-8
            record["native_definition"] = definition
            save()
            result = str(call("AccuracyProductionRun", str(folder / "production.tsv"), str(folder / args.production_case)))
            record["production_result"] = result
            record["precision_target_met"] = result.startswith("AUDITED;") and ";target_met=True;" in result
            print(result, flush=True)
            assert not result.startswith("FAIL;"), result
        elif not args.compile_only:
            if args.quadrature_only:
                result = str(call("AccuracyQGuardProbe"))
                assert result.startswith("PASS;"), result
                record["quadrature_guard"] = result
                print(result, flush=True)
            if args.fan_only or args.fan_geometry_only:
                result = str(call("AccuracyFanGeometryProbe"))
                assert result.startswith("PASS;"), result
                record["fan_geometry"] = result
                print(result, flush=True)
                save()
            if args.fan_geometry_only:
                cases = []
            elif args.fan_only:
                cases = [("fan_disabled", 256, 0, False), ("fan_enabled", 256, 0, False)]
            elif args.quadrature_only:
                cases = [("q8_complete", 64, 0, False), ("q8_injected_failure", 64, 0, False)]
            else:
                cases = [("disabled", 64, 0, False), ("injected_failure", 256, 2, True)] if args.failure_only else [("disabled", 64, 0, False), ("continued", 256, 2, False)]
            for label, cap, rounds, inject_failure in cases:
                trace = folder / (label + ".tsv")
                print("START", label, flush=True)
                started = time.perf_counter()
                result = str(call("AccuracyProbeRun", cap, rounds, str(trace), inject_failure, label == "fan_enabled", args.fan_only, 8 if args.quadrature_only else 0, label == "q8_injected_failure"))
                assert result.startswith("PASS;"), result
                record["runs"][label] = {"result": result, "seconds": time.perf_counter() - started}
                print("END", label, result.split(";upper_identity=", 1)[0], flush=True)
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
            elif args.fan_only:
                values = lambda label: dict(item.split("=", 1) for item in record["runs"][label]["result"].split(";")[1:])
                before, after = values("fan_disabled"), values("fan_enabled")
                assert (before["upper"], before["upper_elements"], before["upper_identity"]) == (after["upper"], after["upper_elements"], after["upper_identity"]), "Independent upper witness changed in lower fan trial"
                logs = [read_log(p) for p in (folder / "RPFEM_logs").glob("*.tsv")]
                assert any("angular_fan_candidate" in log for log in logs)
                assert any("lower_angular_fan_end" in log for log in logs), "Lower candidate root/audit did not finish"
                record["fan_checks"] = "PASS;geometry_payload_integrity;lower_original_gate;independent_upper_exactly_retained"
            elif args.quadrature_only:
                values = lambda label: dict(item.split("=", 1) for item in record["runs"][label]["result"].split(";")[1:])
                complete, failed = values("q8_complete"), values("q8_injected_failure")
                assert complete["lower_q"] == complete["upper_q"] == "8"
                assert complete["lower_root"] == complete["upper_root"] == "True"
                assert failed["lower_q"] == "2" and failed["upper_q"] == "8" and failed["incomplete"] == "True"
                logs = [read_log(p) for p in (folder / "RPFEM_logs").glob("*.tsv")]
                assert any("own_quadrature_accept" in log for log in logs)
                assert any("own_quadrature_rejected" in log and "CONIC_NOT_CONVERGED" in log for log in logs)
                record["quadrature_checks"] = "PASS;independent_q8_original_roots_and_fields;failure_keeps_original_lower_q;precision_target_guard"
            elif not args.fan_geometry_only:
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
    print(json.dumps({key: value for key, value in record.items() if key != "runs"}, ensure_ascii=False), flush=True)


if __name__ == "__main__":
    main()
