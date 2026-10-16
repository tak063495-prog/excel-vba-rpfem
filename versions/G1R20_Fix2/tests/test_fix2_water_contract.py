"""Portable guard contracts for persistent failed seep output recovery."""
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
OLD = ROOT.parent / "G1R20_Fix1/src/vba"


def procedure(text, name):
    match = re.search(r"^(?:Public|Private) (?:Sub|Function) " + re.escape(name) +
                      r"\b.*?^End (?:Sub|Function)\s*$", text, re.M | re.S)
    if match is None:
        raise AssertionError(name)
    return match.group(0).replace("\r", "")


class WaterPersistentContract(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.text = (ROOT / "src/vba/RPX_Water.bas").read_text(encoding="utf8")

    def test_pending_is_journaled_before_any_sheet_write(self):
        begin = procedure(self.text, "RPX_WaterResultWriteBegin")
        self.assertLess(begin.index("wResultWritePending = True"), begin.index("WaterResultWriteMark(WRITE_MARKER_PENDING)"))
        seep = (ROOT / "src/vba/RPX_Seep.bas").read_text(encoding="utf8")
        writer = procedure(seep, "RPX_SeepWriteResults")
        self.assertLess(writer.index("Call RPX_WaterResultWriteBegin"), writer.index('Set ws = RPX_Sheet("浸透結果")'))

    def test_journal_is_workbook_local_and_outside_result_sheet(self):
        getter = procedure(self.text, "WaterResultWriteProperty")
        setter = procedure(self.text, "WaterResultWriteMark")
        self.assertIn("ThisWorkbook.CustomDocumentProperties", getter)
        self.assertIn("ThisWorkbook.CustomDocumentProperties.Add", setter)
        self.assertNotIn("ActiveWorkbook", getter + setter)
        self.assertNotIn(".Range(", getter + setter)
        self.assertIn("item.Type <> 4", setter)
        self.assertIn("SEEP_RESULT_WRITE_STATE_NOT_RECORDED", setter)

    def test_only_completed_writer_clears_in_memory_gate_after_durable_commit(self):
        commit = procedure(self.text, "RPX_WaterResultWriteCommit")
        self.assertLess(commit.index("WaterResultWriteMark(WRITE_MARKER_COMMITTED & identifier)"),
                        commit.index("wResultWritePending = False"))
        self.assertEqual(len(re.findall(r"^\s*wResultWritePending = False\s*$", self.text, re.M)), 1)

    def test_durable_pending_and_wrong_id_fail_closed_before_warm_cache(self):
        assertion = procedure(self.text, "WaterResultWriteAssert")
        self.assertIn('If item Is Nothing Then Exit Sub', assertion)
        self.assertIn('If state = WRITE_MARKER_PENDING Then err.Raise', assertion)
        self.assertIn('state <> WRITE_MARKER_COMMITTED & identifier', assertion)
        self.assertIn('INVALID_SEEP_RESULT_WRITE_STATE', assertion)
        load = procedure(self.text, "LoadSeepResult")
        self.assertLess(load.index("Call WaterResultWriteAssert\n"), load.index("Set ws ="))
        self.assertLess(load.index("Call WaterResultWriteAssert(stamp)"), load.index("If stamp = wStamp"))

    def test_common_mesh_guard_allows_suspended_producer_and_dry_analysis(self):
        common = procedure(self.text, "RPX_WaterCommonLevel")
        self.assertLess(common.index("If RPX_WaterSuspend Or Not CouplingOnSafe() Then Exit Function"),
                        common.index("Call WaterResultWriteAssert"))
        self.assertLess(common.index("Call WaterResultWriteAssert"), common.index("If Not WCCommonInputLevel"))

    def test_cache_input_reset_cannot_clear_persistent_marker(self):
        for name in ("RPX_WaterCacheInvalidate", "RPX_WaterCommonInputBegin", "RPX_WaterCommonInputEnd"):
            body = procedure(self.text, name)
            self.assertNotIn("WaterResultWriteMark", body)
            self.assertNotIn("wResultWritePending = False", body)

    def test_physical_interpolation_and_load_equations_unchanged(self):
        old = (OLD / "RPX_Water.bas").read_text(encoding="utf8")
        for name in ("RPX_WaterApply", "RPX_WaterHeadAt", "BuildGrid", "bary"):
            self.assertEqual(procedure(self.text, name), procedure(old, name), name)


if __name__ == "__main__":
    unittest.main()
