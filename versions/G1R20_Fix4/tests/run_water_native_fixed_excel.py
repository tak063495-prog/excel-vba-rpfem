"""Owned final-book D08/q8 fixed-point regression; no imported fields or roots.

Run only after the shared Excel owner has released it.  The delivered product
modules are never replaced: only two test modules are added to a disposable copy.
"""
from pathlib import Path
import argparse
import gc
import hashlib
import importlib.util
import json
import os
import shutil
import threading
import time
import traceback

for _name in ("OMP_NUM_THREADS", "OPENBLAS_NUM_THREADS", "MKL_NUM_THREADS", "NUMEXPR_NUM_THREADS"):
    os.environ[_name] = "1"

import numpy as np
import psutil
import win32com.client
import win32con
import win32gui
import win32process

ROOT = Path(__file__).resolve().parents[1]
REPO = ROOT.parents[1]
DELIVERY = ROOT / "delivery"
EXPECTED_BOOK = json.loads((ROOT / "SOURCE_MANIFEST.json").read_text(encoding="utf8"))["assets"]["delivery/RPFEM_20261017_G1R20_Fix4.xlsm"]["sha256"]
INPUT = REPO / "docs/benchmarks/G1R17_native26_20261013/data/inputs/D08A.json"
MICRO = REPO / "validation/G1R19_WaterCommon_excel_20261016/scripts/WaterMicroProbe.bas"
AUDITOR = REPO / "versions/G1R20/tests/audit_geometry_replay.py"
CSR_AUDITOR = REPO / "versions/G1R18/tests/test_accuracy_native_excel.py"
NUMERICAL_TREE = REPO / "validation/G1R19_python26_20261016/python"


def sha(path):
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def module_text(path):
    return "\r\n".join(line for line in Path(path).read_text(encoding="utf8").splitlines()
                         if not line.startswith("Attribute VB_"))


def load_module(name, path):
    spec = importlib.util.spec_from_file_location(name, path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def independent_point(prefix, side, factor):
    """Reassemble original coupled q8 equations and evaluate the new native x."""
    import struct
    replay = load_module("fix1_water_literal_csr", CSR_AUDITOR)
    auditor = load_module("fix1_water_same_geometry", AUDITOR)
    proof = Path(str(prefix) + "_accepted_original_primal.bin")
    literal = replay.replay(proof)
    assert literal["fs"] == factor and literal["q"] == 8
    geometry_path = Path(str(prefix) + "_geometry.bin")
    xy, cells, materials, body, boundaries, scale = auditor.geometry(geometry_path)
    auditor.models.install_model_builder()
    canonical = auditor.models._record("D08A")
    mesh = auditor.core.mesh_from_arrays(xy, cells[:, :3], cells[:, 3], canonical)
    np.testing.assert_array_equal(materials, mesh.materials)
    body_difference = float(np.max(np.abs(body-mesh.body), initial=0.))
    np.testing.assert_allclose(body, mesh.body, rtol=0, atol=1e-12)
    np.testing.assert_array_equal(cells[:, 4], mesh.rigid_id)
    assert scale == mesh.stress_scale and boundaries.keys() == mesh.boundary.keys()
    for edge, expected in mesh.boundary.items():
        got = boundaries[edge]
        assert (got["kind"], got["rigid"]) == (expected["kind"], expected["rigid"])
        np.testing.assert_array_equal(got["loads"], expected["loads"])
    points = xy[cells[:, :3]]
    u, v = points[:, 1] - points[:, 0], points[:, 2] - points[:, 0]
    areas = (u[:, 0] * v[:, 1] - u[:, 1] * v[:, 0]) * .5
    assert np.all(areas > 0)
    np.testing.assert_allclose(areas.sum(), 158., rtol=0, atol=1e-9)
    np.testing.assert_allclose(areas @ body.sum(axis=2), [0., -2571.4], rtol=0, atol=1e-8)
    assert not np.any((points[:, :, 1].min(axis=1) < 3.) & (points[:, :, 1].max(axis=1) > 3.))
    edges = {}
    for tri in cells[:, :3]:
        for a, b in zip(tri, np.roll(tri, -1)):
            edge = tuple(sorted((int(a), int(b))))
            edges[edge] = edges.get(edge, 0) + 1
    water_edges = [edge for edge, count in edges.items()
                   if count == 2 and np.all(xy[list(edge), 1] == 3.)]
    assert water_edges and not any(edge in boundaries for edge in water_edges)
    prepared = auditor.core.Prepared(mesh, factor, side, 8, reduction="none")
    exported = auditor.original_program(Path(str(prefix) + "_original_assembled.bin"))
    comparisons = {key: auditor.difference(left, right) for key, left, right in
                   zip(("E", "b", "G", "h", "cost"), exported[:4] + (exported[5],),
                       prepared.original[:4] + (prepared.original[5],))}
    comparisons["cone_dimensions"] = {"pass_": bool(np.array_equal(exported[4], prepared.original[4]))}
    assert all(row["pass_"] for row in comparisons.values()), comparisons
    raw = proof.read_bytes()
    n = struct.unpack_from("<i", raw, 8)[0]
    x = np.frombuffer(raw, dtype="<f8", count=n, offset=len(raw)-8*n).copy()
    before = hashlib.sha256(x.tobytes()).hexdigest()
    evaluated = auditor.core.evaluate(prepared, x)
    assert before == hashlib.sha256(x.tobytes()).hexdigest()
    assert evaluated["physical_pass"] and evaluated["witness"], evaluated
    return dict(status="FIXED_POINT_PARITY_PASS", side=side, factor=factor,
                q=8, nodes=len(xy), elements=len(cells), area=float(areas.sum()),
                reference_force=(areas @ body.sum(axis=2)).tolist(),
                internal_waterline_edges=len(water_edges), waterline_physical_boundary=False,
                canonical_material_boundary_exact=True, body_max_absolute_difference=body_difference,
                body_equivalence_tolerance=1e-12, comparisons=comparisons,
                literal_original_csr=literal, independent_original_q8=evaluated,
                field_unchanged=True, root_complete=False,
                scope="Fresh fixed-point primal only; no complete root or adaptive-run claim")


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--seconds", type=float, default=240.)
    parser.add_argument("--solve-seconds", type=float, default=20.)
    args = parser.parse_args()
    if not 30 <= args.seconds <= 300 or not .1 <= args.solve_seconds <= 30:
        parser.error("bounded test budget required")
    prior = [p.pid for p in psutil.process_iter() if p.name().upper() == "EXCEL.EXE"]
    if prior:
        raise RuntimeError("Another Excel exists: " + repr(prior))
    source_book = DELIVERY / "RPFEM_20261017_G1R20_Fix4.xlsm"
    source_dll = DELIVERY / "rpfem_native_x64.dll"
    assert sha(source_book) == EXPECTED_BOOK, "Final delivery workbook changed"
    probes = [ROOT / "tests/WaterNativeFixedProbe.bas", MICRO]
    product = list((ROOT / "src/vba").glob("*.bas"))
    assert len(product) == 63
    paths = product + list((ROOT / "import").glob("*.bas")) + [ROOT / "src/workbook/ThisWorkbook.code.txt", ROOT / "import/ThisWorkbook.code.txt"] + probes
    paths += [source_book, source_dll, ROOT / "SOURCE_MANIFEST.json", INPUT, AUDITOR,
              CSR_AUDITOR, Path(__file__)] + list(NUMERICAL_TREE.rglob("*.py"))
    paths = sorted(set(path.resolve() for path in paths))
    before = {str(path): sha(path) for path in paths}
    out = ROOT / "results" / ("water_native_fixed_" + str(time.time_ns()))
    out.mkdir(parents=True)
    bookpath = out / "WaterNativeFixed.xlsm"
    shutil.copyfile(source_book, bookpath)
    shutil.copyfile(source_dll, out / source_dll.name)
    for path in probes + [Path(__file__), INPUT]:
        shutil.copyfile(path, out / path.name)
    rec = dict(status="PREPARING", case="D08A", target=64, q=8, points={}, dialogs=[],
               expected_product_modules=63, product_modules_replaced=False,
               input_geometry_field_root_imported=False, full_root_claimed=False,
               source_sha256_before=before, prior_excel_pids=prior,
               python_owner=dict(pid=psutil.Process().pid, create_time=psutil.Process().create_time(),
                                 cmdline=psutil.Process().cmdline()),
               hard_budget_seconds=args.seconds, per_solve_seconds=args.solve_seconds,
               started_epoch=time.time())
    def save():
        (out / "summary.json").write_text(json.dumps(rec, indent=2, ensure_ascii=False, allow_nan=False) + "\n", encoding="utf8")
    save()
    app = book = component = control = ws = None
    stop = threading.Event()
    try:
        app = win32com.client.DispatchEx("Excel.Application")
        pid = win32process.GetWindowThreadProcessId(app.Hwnd)[1]
        owner = psutil.Process(pid)
        birth, command = owner.create_time(), owner.cmdline()
        rec["excel_owner"] = dict(pid=pid, create_time=birth, cmdline=command, exe=owner.exe())
        save()
        def watch():
            deadline = time.monotonic() + args.seconds
            while not stop.wait(.1):
                if time.monotonic() > deadline:
                    try:
                        process = psutil.Process(pid)
                        if process.create_time() == birth and process.cmdline() == command:
                            process.kill()
                            rec["watchdog"] = "EXACT_OWNED_DEADLINE_STOP"
                            save()
                    except psutil.NoSuchProcess:
                        pass
                    return
                def visit(window, unused):
                    if win32process.GetWindowThreadProcessId(window)[1] != pid or win32gui.GetClassName(window) != "#32770":
                        return
                    children = []
                    win32gui.EnumChildWindows(window, lambda c, _: children.append((c, win32gui.GetWindowText(c))), None)
                    words = [value for _, value in children]
                    if words not in rec["dialogs"]:
                        rec["dialogs"].append(words)
                        save()
                    for child, label in children:
                        if label in ("OK", "&OK") or "&End" in label or "終了" in label:
                            win32gui.PostMessage(child, win32con.BM_CLICK, 0, 0)
                try:
                    win32gui.EnumWindows(visit, None)
                except Exception as monitor_error:
                    code = getattr(monitor_error, "winerror", monitor_error.args[0] if monitor_error.args else None)
                    key = "transient_window_1400" if code == 1400 else "window_monitor_errors"
                    errors = rec.setdefault(key, [])
                    if repr(monitor_error) not in errors:
                        errors.append(repr(monitor_error))
        watcher = threading.Thread(target=watch, daemon=True)
        watcher.start()
        app.Visible = False
        app.DisplayAlerts = False
        app.EnableEvents = False
        app.MultiThreadedCalculation.Enabled = False
        book = app.Workbooks.Open(str(bookpath), UpdateLinks=0)
        loaded_before = {c.Name: c.CodeModule.Lines(1, c.CodeModule.CountOfLines)
                         for c in book.VBProject.VBComponents if c.Type == 1}
        rec["loaded_product_sha256_before"] = {
            name: hashlib.sha256(text.encode("utf8")).hexdigest() for name, text in loaded_before.items()}
        assert set(rec["loaded_product_sha256_before"]) == {p.stem for p in product}
        for path in probes:
            component = book.VBProject.VBComponents.Add(1)
            component.Name = path.stem
            component.CodeModule.AddFromString(module_text(path))
        component.CodeModule.CodePane.Show()
        control = app.VBE.CommandBars.FindControl(1, 578)
        if control.Enabled:
            control.Execute()
            time.sleep(.2)
        if control.Enabled:
            pane = app.VBE.ActiveCodePane
            line = int(pane.GetSelection()[0])
            raise RuntimeError("VBA compile failed " + pane.CodeModule.Name + ":" + pane.CodeModule.Lines(max(1, line-2), 5))
        loaded_compiled = {c.Name: c.CodeModule.Lines(1, c.CodeModule.CountOfLines)
                           for c in book.VBProject.VBComponents if c.Type == 1 and c.Name in loaded_before}
        for phase, modules in (("before_test_probes", loaded_before), ("after_test_compile", loaded_compiled)):
            directory = out / "loaded_product_snapshot" / phase
            directory.mkdir(parents=True)
            for name, text in modules.items():
                (directory / (name + ".bas")).write_bytes(text.encode("utf8"))
        assert all(loaded_before[name].lower() == loaded_compiled[name].lower() for name in loaded_before), "Probe compile changed product beyond identifier case"
        rec["probe_compile_product_changes_case_only"] = True
        rec["probe_compile_case_changed_modules"] = [name for name in loaded_before if loaded_before[name] != loaded_compiled[name]]
        rec["loaded_product_sha256_after_compile"] = {
            name: hashlib.sha256(text.encode("utf8")).hexdigest() for name, text in loaded_compiled.items()}
        rec["compile"] = "PASS"
        save()
        app.VBE.MainWindow.Visible = False
        def call(name, *values):
            return app.Run("'" + book.Name + "'!" + name, *values)
        call("RPX_Setup")
        data = json.loads(INPUT.read_text(encoding="utf8"))
        for sheet, ranges in [("要素定義", [("A", "C", "points", 7), ("E", "H", "regions", 7), ("J", "T", "faces", 7)]),
                              ("材料データ", [("A", "E", "materials", 5)])]:
            ws = book.Worksheets(sheet)
            ws.Range("A7:T1006" if sheet == "要素定義" else "A5:Q1004").ClearContents()
            for first, last, key, row in ranges:
                values = tuple(tuple(value) for value in data[key])
                ws.Range(f"{first}{row}:{last}{row+len(values)-1}").Value2 = values
        ws = book.Worksheets("設定")
        def setting(key, value):
            rows = ws.Range("D4:D999").Value2
            ids = [index+4 for index, row in enumerate(rows) if row[0] == key]
            if not ids:
                empty = [index+4 for index, row in enumerate(ws.Range("B4:D999").Value2)
                         if row[0] is None and row[2] is None]
                if not empty:
                    raise ValueError("No row for canonical setting " + key)
                ids = [empty[0]]
                ws.Cells(ids[0], 4).Value2 = key
            if len(ids) != 1:
                raise ValueError("Duplicate setting " + key)
            ws.Cells(ids[0], 2).Value2 = value
        for key, value in data["settings"].items():
            setting(key, value)
        call("RPX_Speed2pctPreset")
        for key, value in dict(TARGET=64, QUADRATURE=8, ADAPT=0, DIRECT_DLL_SECONDS=args.solve_seconds,
                               ACCURACY_DLL_DUAL_SECONDS=10, ACCURACY_DLL_SECONDS=args.solve_seconds,
                               ACCURACY_DLL_CAPTURE=1, DIAGNOSTIC_MODE=1).items():
            setting(key, value)
        rec["settings"] = {str(row[2]): row[0] for row in ws.Range("B4:D999").Value2 if row[2]}
        ws = None
        rec["prescribed_result"] = call("WaterMicroPrepareResult")
        save()
        assert rec["prescribed_result"].startswith("PASS"), rec["prescribed_result"]
        for side, factor in (("lower", 1.), ("upper", 2.)):
            prefix = out / side
            started = time.perf_counter()
            result = list(call("WaterNativeFixedPoint", side, factor, str(prefix)))
            rec["points"][side] = dict(result=result, elapsed_seconds=time.perf_counter()-started)
            save()
            assert result[0] == "PASS", result
        rec["loaded_product_sha256_after"] = {
            c.Name: hashlib.sha256(c.CodeModule.Lines(1, c.CodeModule.CountOfLines).encode("utf8")).hexdigest()
            for c in book.VBProject.VBComponents if c.Type == 1 and c.Name in rec["loaded_product_sha256_before"]}
        assert rec["loaded_product_sha256_after_compile"] == rec["loaded_product_sha256_after"]
        mappings = [row.path for row in owner.memory_maps() if row.path.lower().endswith(".dll") and "rpfem" in row.path.lower()]
        rec["native_mapped_paths"] = mappings
        expected = (out / "rpfem_native_x64.dll").resolve()
        assert mappings and all(Path(path).resolve() == expected for path in mappings), mappings
        assert sha(expected) == sha(source_dll)
        rec["single_dll_mapping_expected_path_hash"] = True
        book.Save()
        rec["status"] = "VBA_FIXED_POINTS_PASS"
        save()
    except BaseException as error:
        rec.update(status="FAIL", error=repr(error), traceback=traceback.format_exc())
        save()
    finally:
        try:
            if book is not None:
                book.Close(False)
            if app is not None:
                app.Quit()
        except Exception as error:
            rec["close_error"] = repr(error)
        component = control = ws = book = app = None
        gc.collect()
        stop.set()
        if "watcher" in locals():
            watcher.join(2)
        if "pid" in locals():
            for _ in range(50):
                if not psutil.pid_exists(pid) or psutil.Process(pid).create_time() != birth:
                    break
                time.sleep(.1)
            rec["owned_excel_exited"] = not psutil.pid_exists(pid) or psutil.Process(pid).create_time() != birth
        rec["source_sha256_after"] = {str(path): sha(path) for path in paths}
        rec["source_unchanged"] = before == rec["source_sha256_after"]
        rec["excel_finished_epoch"] = time.time()
        save()
    if rec["status"] == "VBA_FIXED_POINTS_PASS":
        try:
            for side, factor in (("lower", 1.), ("upper", 2.)):
                rec["points"][side]["independent_audit"] = independent_point(out / side, side, factor)
                save()
            rec["same_geometry_lower_upper"] = sha(out / "lower_geometry.bin") == sha(out / "upper_geometry.bin")
            assert rec["same_geometry_lower_upper"]
            assert rec["owned_excel_exited"] and rec["source_unchanged"] and not rec["dialogs"] and not rec.get("window_monitor_errors")
            rec["status"] = "PASS"
        except BaseException as error:
            rec.update(status="FAIL", audit_error=repr(error), audit_traceback=traceback.format_exc())
    rec["source_sha256_after_audit"] = {str(path): sha(path) for path in paths}
    rec["source_unchanged"] = before == rec["source_sha256_after_audit"]
    if not rec["source_unchanged"]:
        rec["status"] = "FAIL"
    rec["finished_epoch"] = time.time()
    rec["elapsed_seconds"] = rec["finished_epoch"] - rec["started_epoch"]
    save()
    print(json.dumps(dict(output=str(out), status=rec["status"], points=rec["points"],
                         owned_excel_exited=rec.get("owned_excel_exited"), source_unchanged=rec["source_unchanged"]),
                     ensure_ascii=False, allow_nan=False), flush=True)
    return 0 if rec["status"] == "PASS" else 2


if __name__ == "__main__":
    raise SystemExit(main())
