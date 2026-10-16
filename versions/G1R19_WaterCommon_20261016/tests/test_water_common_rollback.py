"""Rollback failure contract/static checks; actual VBA probes are separate."""
from pathlib import Path
import unittest

ROOT=Path(__file__).resolve().parents[1]
SOURCE=(ROOT/'src/vba/RPX_Mesh.bas').read_text(encoding='utf8')


def outcome(original, restore=None, cancelled=False):
    if restore is not None:
        code=18 if original[0]==18 or restore[0]==18 else 5
        return code,dict(kind='ROLLBACK_FAILED',original=original,restore=restore),True
    token=original[1].split(' ')[0]
    if original[0]==18 or cancelled or token in ('INPUT_MISMATCH','ANALYSIS_CANCELLED','WATER_COMMON_SOURCE_CHANGED'):
        return original[0],original[1],True
    return 0,'VERIFIED_LEGACY_FALLBACK',False


class RollbackTests(unittest.TestCase):
    def test_verified_restore_only_allows_ordinary_fallback(self):
        self.assertEqual(outcome((5,'ordinary geometry rejection')),(0,'VERIFIED_LEGACY_FALLBACK',False))
        failed=SOURCE.split('Failed:\n    errorNumber = err.number',1)[1].split('End Function',1)[0]
        self.assertLess(failed.index('WCRestoreMesh('),failed.index('unchanged legacy fallback'))
        self.assertLess(failed.index('RPX_AssertInputs "water_common_rollback_verified"'),failed.index('unchanged legacy fallback'))

    def test_restore_failure_is_terminal_and_preserves_both_errors(self):
        for restore in ((5,'topology failure'),(7,'allocation failure'),(5,'face mismatch')):
            code,detail,cancelled=outcome((5,'original clipping failure'),restore)
            self.assertEqual(code,5);self.assertTrue(cancelled)
            self.assertEqual(detail['original'],(5,'original clipping failure'))
            self.assertEqual(detail['restore'],restore)
        for token in ('original_number=','original_source=','original_message=','restore_number=','restore_source=','restore_message='):
            self.assertIn(token,SOURCE)
        self.assertIn('err.Raise terminalNumber, "RPX_Mesh", failureMessage',SOURCE)

    def test_original_cancel_always_propagates(self):
        self.assertEqual(outcome((18,'cancel'))[0],18)
        self.assertEqual(outcome((18,'cancel'),(5,'restore failed'))[0],18)
        self.assertIn('If errorNumber = 18 Or restoreNumber = 18 Then terminalNumber = 18',SOURCE)

    def test_restore_cancel_and_input_mismatch_always_propagate(self):
        self.assertEqual(outcome((5,'ordinary'),(18,'cancel in restore'))[0],18)
        self.assertEqual(outcome((5,'INPUT_MISMATCH test'))[0],5)
        self.assertEqual(outcome((5,'WATER_COMMON_SOURCE_CHANGED test'))[0],5)
        self.assertEqual(outcome((5,'ordinary'),cancelled=True)[0],5)

    def test_restore_uses_separate_handler_and_boundary_verification(self):
        helper=SOURCE.split('Private Function WCRestoreMesh(',1)[1].split('End Function',1)[0]
        self.assertIn('On Error GoTo Failed',helper)
        self.assertNotIn('On Error Resume Next',helper)
        self.assertIn('restoreNumber = err.number: restoreSource = err.source: restoreText = err.description',helper)
        self.assertIn('If Not WCBoundaryEqual(oldBoundary)',helper)
        self.assertLess(helper.index('Call RPX_Topology'),helper.index('Call RPX_MapFaces'))

    def test_injections_use_existing_armed_test_mechanism_only(self):
        for stage in ('water_common_after_commit','water_common_rollback_arrays','water_common_rollback_topology','water_common_rollback_faces'):
            self.assertIn('RPX_TestHook "'+stage+'"',SOURCE)
        self.assertNotIn('TEST_INJECTION',SOURCE)
        self.assertLess(max(map(len,SOURCE.splitlines())),1023)


if __name__=='__main__':
    unittest.main()
