"""Source binding contract checks; no claim of executing VBA/Excel.

The small typed-state model tests the required identity semantics. Static
checks establish where the VBA hook reads, binds, rejects, and resets them.
WaterCommonSourceProbe.bas separately exercises the actual VBA in a test copy.
"""
import copy
import re
import struct
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
WATER = (ROOT/'src/vba/RPX_Water.bas').read_text(encoding='utf8')
GUARD = (ROOT/'src/vba/RPX_Guard.bas').read_text(encoding='utf8')
REPO = ROOT.parents[1]


def cell(kind, value):
    payload = struct.pack('<d', value).hex().upper() if kind == 5 else str(value)
    return f'{kind}:{len(payload)}:{payload}'


def signature(context):
    return tuple((key, cell(*value)) for key,value in sorted(context.items()))


class Pin:
    def begin(self, context, epoch, eligible=True):
        self.end()
        self.scope, self.epoch = True, epoch
        self.bound = eligible
        if eligible:
            self.source = signature(context)

    def end(self):
        self.scope = self.bound = False
        self.source = None

    def check(self, context, epoch):
        if self.scope and self.bound and (epoch != self.epoch or signature(context) != self.source):
            raise ValueError('WATER_COMMON_SOURCE_CHANGED')

    def common(self):
        return not self.scope or self.bound


class SourceTests(unittest.TestCase):
    def context(self):
        return {'gamma':(5,9.81), 'coupling':(5,1.), 'suction':(5,0.), 'mode':(8,'static'),
                'table_x':(5,0.), 'table_y':(5,3.), 'stamp':(8,'fresh'), 'head':(5,3.),
                'node_x':(5,0.), 'triangle':(5,1.), 'converged':(11,True)}

    def test_exact_double_type_and_length_delimiting(self):
        self.assertNotEqual(cell(5,0.), cell(5,-0.))
        self.assertNotEqual(cell(5,3.), cell(8,'3'))
        self.assertNotEqual(cell(5,3.), cell(5,3.+2**-50))
        self.assertNotEqual(cell(8,'1:a|8:1:b'), cell(8,'a|8:1:b'))
        for token in ('VarType(value)', 'RPX_ExactDouble(CDbl(value))', 'Len(text)', 'cells.formula'):
            self.assertIn(token,WATER)

    def test_every_bound_water_source_change_rejected(self):
        original=self.context(); pin=Pin();pin.begin(original,7)
        for key in original:
            changed=copy.deepcopy(original);kind,value=changed[key]
            changed[key]=(kind, value+.01 if kind==5 else not value if kind==11 else value+':changed')
            with self.subTest(key=key), self.assertRaisesRegex(ValueError,'SOURCE_CHANGED'):
                pin.check(changed,7)
        for address in ('Z1:Z7','B3:F4','M7:N206','K2:K10','A2','R2'):
            self.assertIn('Range("'+address+'")',WATER)

    def test_epoch_and_missing_source_rejected(self):
        pin=Pin();source=self.context();pin.begin(source,7)
        with self.assertRaises(ValueError):pin.check(source,8)
        del source['head']
        with self.assertRaises(ValueError):pin.check(source,7)
        self.assertIn('commonInputEpoch <> RPX_InputEpoch',WATER)
        self.assertIn('WATER_COMMON_SOURCE_SHEET_MISSING',WATER)

    def test_end_newrun_and_preview_reset(self):
        pin=Pin();source=self.context();pin.begin(source,7);pin.end()
        changed=dict(source,gamma=(5,10.));pin.check(changed,8)
        self.assertTrue(pin.common())  # Preview outside a frozen analysis.
        pin.begin(changed,8);pin.check(changed,8)
        self.assertIn('commonInputScope = False: commonInputBound = False',WATER)
        self.assertIn('wStamp = ""',WATER)

    def test_initially_ineligible_run_cannot_engage_later(self):
        pin=Pin();pin.begin(self.context(),7,eligible=False)
        self.assertFalse(pin.common())
        self.assertIn('If commonInputScope And Not commonInputBound Then Exit Function',WATER)

    def test_no_recursive_assert_or_silent_fallback(self):
        public=WATER.split('Public Function RPX_WaterCommonLevel',1)[1].split('End Function',1)[0]
        private=WATER.split('Private Function WCCommonInputLevel',1)[1].split('End Function',1)[0]
        check=WATER.split('Public Sub RPX_WaterCommonInputAssert()',1)[1].split('End Sub',1)[0]
        self.assertNotIn('InputAssert',private)
        self.assertNotIn('On Error GoTo Ineligible',public)
        self.assertLess(public.index('InputAssert'),public.index('WCCommonInputLevel'))
        self.assertIn('XCancel = True: RPX_ResultCurrent = False: RPX_MeshCurrent = False',check)
        self.assertIn('err.Raise 5, "RPX_Water", "WATER_COMMON_SOURCE_CHANGED',check)
        self.assertIn('If commonInputScope And commonEngaged And Not commonMode Then',WATER)
        self.assertIn('WATER_COMMON_LOST_CONFORMITY_OR_ELIGIBILITY',WATER)
        self.assertIn('If commonInputScope And commonMode Then commonEngaged = True',WATER)
        self.assertIn('commonEngaged = False',WATER)

    def test_standard_begin_stage_final_and_end_hooks(self):
        start=GUARD.split('Public Sub RPX_InputBegin()',1)[1].split('End Sub',1)[0]
        check=GUARD.split('Public Sub RPX_AssertInputs(',1)[1].split('End Sub',1)[0]
        end=GUARD.split('Public Sub RPX_InputEnd()',1)[1].split('End Sub',1)[0]
        self.assertLess(start.index('RPX_InputEpoch = RPX_InputEpoch + 1'),start.index('Call RPX_WaterCommonInputBegin'))
        self.assertLess(check.index('Call RPX_WaterCommonInputAssert'),check.index('If RPX_InputsChanged()'))
        self.assertIn('Call RPX_WaterCommonInputEnd',end)
        control=(REPO/'versions/G1R19/src/vba/RPX_Control.bas').read_text(encoding='utf8')
        self.assertIn('RPX_AssertInputs "result_commit"',control)

    def test_legacy_guard_changes_are_hooks_only(self):
        original=(REPO/'versions/G1R19/src/vba/RPX_Guard.bas').read_text(encoding='utf8')
        without_hooks='\n'.join(line for line in GUARD.splitlines() if 'Call RPX_WaterCommonInput' not in line)+'\n'
        self.assertEqual(original,without_hooks)
        self.assertLess(max(map(len,WATER.splitlines())),1023)


if __name__=='__main__':
    unittest.main()
