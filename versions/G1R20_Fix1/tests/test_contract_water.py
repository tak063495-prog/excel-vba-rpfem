"""Portable water-result lifecycle contracts; no Excel or numeric solver."""
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT.parent / 'G1R20'


def source(name, base=ROOT):
    return (base / 'src/vba' / name).read_text(encoding='utf8')


def procedure(text, name):
    pattern = (r'(?ms)^(?:Public|Private) (?:Sub|Function) ' + re.escape(name) +
               r'\b.*?^End (?:Sub|Function)\s*$')
    match = re.search(pattern, text)
    if not match:
        raise AssertionError('Missing procedure ' + name)
    return match.group()


class WaterResultLifecycle(unittest.TestCase):
    def test_all_mode_input_boundaries_clear_arrays_before_eligibility(self):
        text = source('RPX_Water.bas')
        begin = procedure(text, 'RPX_WaterCommonInputBegin')
        self.assertLess(begin.index('Call RPX_WaterCacheInvalidate'), begin.index('If Not WCCommonInputLevel'))
        clear = procedure(text, 'RPX_WaterCacheInvalidate')
        for token in ['wStamp = ""', 'wN = 0', 'wM = 0', 'Erase wHead', 'Erase wXY',
                      'Erase wTri', 'Erase gStart', 'Erase gList']:
            self.assertIn(token, clear)
        self.assertNotRegex(clear, r'wResultWritePending\s*=')
        self.assertNotRegex(procedure(text, 'RPX_WaterCommonInputEnd'), r'wResultWritePending\s*=')

    def test_failed_producer_stays_blocked_until_successful_commit(self):
        text = source('RPX_Water.bas')
        self.assertIn('wResultWritePending = True', procedure(text, 'RPX_WaterResultWriteBegin'))
        load = procedure(text, 'LoadSeepResult')
        self.assertLess(load.index('If wResultWritePending'), load.index('Set ws ='))
        self.assertLess(load.index('If wResultWritePending'), load.index('If stamp = wStamp'))
        assignments = re.findall(r'^\s*wResultWritePending\s*=\s*False\s*$', text, re.M)
        self.assertEqual(len(assignments), 1)
        self.assertIn('wResultWritePending = False', procedure(text, 'RPX_WaterResultWriteCommit'))
        writer = procedure(source('RPX_Seep.bas'), 'RPX_SeepWriteResults')
        self.assertLess(writer.index('Call RPX_WaterResultWriteBegin'), writer.index('Set ws ='))
        self.assertIn('Call RPX_WaterResultWriteBegin', writer.split('Failed:', 1)[1])
        self.assertIn('err.Raise failureNumber, failureSource, failureText', writer)

    def test_id_has_nonwrapping_session_sequence(self):
        text = source('RPX_Seep.bas')
        self.assertIn('Private seepWriteSequence As Long', text)
        build = procedure(text, 'NextSeepResultId')
        self.assertLess(build.index('If seepWriteSequence = 2147483647'),
                        build.index('seepWriteSequence = seepWriteSequence + 1'))
        self.assertIn('SEEP_RESULT_ID_SEQUENCE_EXHAUSTED', build)
        self.assertIn('CStr(seepWriteSequence)', build)
        self.assertEqual(len(re.findall(r'^\s*seepWriteSequence\s*=', text, re.M)), 1)

    def test_eligibility_committed_after_complete_payload(self):
        writer = procedure(source('RPX_Seep.bas'), 'RPX_SeepWriteResults')
        body, failure = writer.split('Failed:', 1)
        self.assertLess(body.index('ws.Range("K2").Value2 = False'), body.index('ws.cells.Clear'))
        self.assertLess(body.index('ws.Range("R2").Resize'), body.index('ws.Range("K9").Value2 = identifier'))
        self.assertLess(body.index('ws.Range("M2").Resize'), body.index('ws.Range("K2").Value2 = SConverged'))
        self.assertLess(body.index('ws.Range("K2").Value2 = SConverged'),
                        body.index('Call RPX_WaterResultWriteCommit(identifier)'))
        self.assertIn('ws.Range("K2").Value2 = False', failure)
        self.assertIn('ws.Range("K9").ClearContents', failure)

    def test_ui_attempt_invalidation_precedes_suppressed_sheet_failure(self):
        old = source('RPX_SeepUI.bas', BASE)
        new = source('RPX_SeepUI.bas')
        helper = procedure(new, 'RPX_SeepInvalidateResult')
        self.assertLess(helper.index('Call RPX_WaterResultWriteBegin'), helper.index('On Error Resume Next'))
        self.assertEqual(new.replace('    Call RPX_WaterResultWriteBegin\n', '', 1), old)

    def test_equations_and_acceptance_procedures_are_unchanged(self):
        water = ['RPX_WaterApply', 'RPX_WaterHeadAt', 'WCCommonInputLevel',
                 'RPX_WaterCommonConforming', 'RPX_WaterCommonInputAssert']
        seep = ['RPX_SeepSolve', 'RPX_SeepFreeSurface', 'RPX_SeepDefaults', 'ElemK', 'TriGeom']
        for name, methods in [('RPX_Water.bas', water), ('RPX_Seep.bas', seep)]:
            for method in methods:
                self.assertEqual(procedure(source(name), method), procedure(source(name, BASE), method))

    def test_cp932_import_pairs_match_sources(self):
        for name in ['RPX_Water.bas', 'RPX_Seep.bas', 'RPX_SeepUI.bas']:
            self.assertEqual(source(name), (ROOT / 'import' / name).read_text(encoding='cp932'))


if __name__ == '__main__':
    unittest.main()
