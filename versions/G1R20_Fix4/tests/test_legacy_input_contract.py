"""Protect the text-reader boundary while retaining all numerical kernels.

The executable Excel cases verify malformed inputs, exact integer handling,
file lock release and the original two-triangle finite-Fs oracle.
"""
from pathlib import Path
import unittest

ROOT=Path(__file__).resolve().parents[1]
BASE=ROOT.parent/'G1R20_Fix3'


class LegacyInputContract(unittest.TestCase):
    def test_mathematical_model_prefix_and_source_pair_are_unchanged(self):
        source=(ROOT/'src/vba/RPX_Model.bas').read_text(encoding='utf8')
        old=(BASE/'src/vba/RPX_Model.bas').read_text(encoding='utf8')
        self.assertEqual(source.split('Private Function LegacyNumber(',1)[0],
                         old.split('Public Sub RPX_ReadModel(',1)[0])
        self.assertEqual(source,(ROOT/'import/RPX_Model.bas').read_text(encoding='cp932'))

    def test_integer_conversion_and_file_cleanup_follow_validation(self):
        source=(ROOT/'src/vba/RPX_Model.bas').read_text(encoding='utf8')
        integer=source.split('Private Function LegacyInteger(',1)[1].split('End Function',1)[0]
        self.assertLess(integer.index('value <> Fix(value)'),integer.index('CLng(value)'))
        reader=source.split('Public Sub RPX_ReadModel(',1)[1]
        self.assertIn('On Error GoTo Failed',reader)
        self.assertIn('Resume FailedCleanup',reader)
        self.assertIn('If opened Then Close #f',reader)
        self.assertIn('err.Raise number, source, message',reader)
        self.assertIn('LEGACY_BOUNDARY_NOT_EXPOSED',reader)
        self.assertIn('Set RConstraints = RPX_Row()',reader)


if __name__=='__main__': unittest.main()
