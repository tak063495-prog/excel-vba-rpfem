"""Isolate whether a journal-only write marks the Excel workbook unsaved."""
from pathlib import Path
import argparse
import importlib.util
import json
import shutil
import time
import traceback

ROOT = Path(__file__).resolve().parents[1]
HARNESS = ROOT / "tests/run_water_restart_excel.py"
PROBE = ROOT / "tests/WaterJournalDirtyFix2Probe.bas"


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument("--execute", action="store_true")
    args = parser.parse_args()
    if not args.execute:
        parser.error("--execute required")
    spec = importlib.util.spec_from_file_location("water_restart_owned_session", HARNESS)
    helper = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(helper)
    prior = [item.pid for item in helper.psutil.process_iter() if item.name().upper() == "EXCEL.EXE"]
    if prior:
        raise RuntimeError("Another Excel exists: " + repr(prior))
    helper.PROBE = PROBE
    paths = list((ROOT / "src/vba").glob("*.bas")) + list((ROOT / "import").glob("*.bas"))
    paths += [HARNESS, PROBE, helper.BRIDGE, Path(__file__), helper.BOOK, helper.DLL]
    before = {str(path.resolve()): helper.sha(path) for path in paths}
    out = ROOT / "results" / ("water_journal_dirty_" + str(time.time_ns()))
    out.mkdir(parents=True)
    for path in [PROBE, HARNESS, Path(__file__)]:
        shutil.copyfile(path, out / path.name)
    book = out / "JournalDirty.xlsm"
    shutil.copyfile(helper.BOOK, book)
    shutil.copyfile(helper.DLL, out / helper.DLL.name)
    rec = dict(status="PREPARING", dialogs=[], source_sha256_before=before, prior_excel_pids=prior, started_epoch=time.time())
    def save():
        (out / "summary.json").write_text(json.dumps(rec, ensure_ascii=False, indent=2) + "\n", encoding="utf8")
    save()
    try:
        with helper.Session(rec, save, out) as excel:
            excel.open(book)
            excel.inject(True)
            result = excel.book.Worksheets("浸透結果")
            cells_before = [result.Range(name).Value2 for name in ("K2", "K9")]
            excel.book.Save()
            saved_before = bool(excel.book.Saved)
            # Call the product entry directly: no test macro or worksheet write
            # is interposed between the two COM Saved observations.
            excel.app.Run("'" + excel.book.Name + "'!RPX_WaterResultWriteBegin")
            saved_after = bool(excel.book.Saved)
            marker = str(excel.book.CustomDocumentProperties("RPX_SeepResultWriteState").Value)
            cells_after = [result.Range(name).Value2 for name in ("K2", "K9")]
            assert cells_before == cells_after
            result = None
            rec["result"] = ["OBSERVED", saved_before, saved_after, marker, "result_K2_K9_unchanged"]
            assert rec["result"][0] == "OBSERVED" and rec["result"][1] is True and rec["result"][3] == "PENDING", rec
            rec["journal_only_marks_unsaved"] = rec["result"][2] is False
            excel.book.Save()
        assert rec["owned_excel_exited"] and not rec["dialogs"] and "monitor_error" not in rec, rec
        rec["custom_properties"] = helper.property_snapshot(book, out / "custom.xml")
        rec["status"] = "PASS" if rec["journal_only_marks_unsaved"] else "NEEDS_EXPLICIT_DIRTY_FLAG"
    except BaseException as error:
        rec.update(status="FAIL", error=repr(error), traceback=traceback.format_exc())
    rec["source_sha256_after"] = {str(path.resolve()): helper.sha(path) for path in paths}
    rec["source_unchanged"] = before == rec["source_sha256_after"]
    rec["finished_epoch"] = time.time()
    save()
    print(json.dumps(dict(output=str(out), status=rec["status"], result=rec.get("result"), owned_excel_exited=rec.get("owned_excel_exited"), source_unchanged=rec["source_unchanged"]), ensure_ascii=False), flush=True)
    return 0 if rec["status"] == "PASS" and rec.get("owned_excel_exited") and rec["source_unchanged"] else 2


if __name__ == "__main__":
    raise SystemExit(main())
