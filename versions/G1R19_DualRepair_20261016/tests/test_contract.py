"""Unchanged original equations plus exact scope/one-shot/free-y contracts."""
from pathlib import Path
import importlib.util,json,re,unittest

ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('dualrepair',ROOT/'build_overlay.py')
overlay=importlib.util.module_from_spec(spec);spec.loader.exec_module(overlay)


def procedure(text,name):
    m=re.search(r'^(?:Public|Private) (Function|Sub) '+re.escape(name)+r'\b',text,re.M)
    if not m:raise ValueError(name)
    last=text.index('\nEnd '+m[1],m.start())+len('\nEnd '+m[1])
    return text[m.start():last]


class Contract(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.old=overlay.BASE.read_text(encoding='utf8')
        cls.new=(ROOT/'src/vba/RPX_AccuracyNative.bas').read_text(encoding='utf8')

    def test_only_declared_patch_changes_source(self):
        self.assertEqual(self.new,overlay.expected())
        text=self.new.removesuffix(overlay.HELPERS)
        for before,after in reversed(overlay.changes()):
            self.assertEqual(text.count(after),1);text=text.replace(after,before,1)
        self.assertEqual(text,self.old)

    def test_original_primal_dual_equations_and_gate_unchanged(self):
        for name in ['CurrentProgram','RPX_AccuracyExport','RPX_AccuracyPrimalAudit',
                     'RPX_AccuracyDualAudit','RPX_AccuracyDualProblem','RPX_AccuracySolve',
                     'RecoverUpperRecenterProgram','NormalizeUpperWork','AddBound','HomogeneousLowerWitness']:
            with self.subTest(name=name):self.assertEqual(procedure(self.old,name),procedure(self.new,name))

    def test_fixed_one_lsmr_call_and_original_dual_audit(self):
        worker=procedure(self.new,'TryOriginalDualLinearRepair')
        self.assertEqual(worker.count('NEqualities('),1)
        self.assertIn('200, 0.000000000000001, 0.000000000000001, 1000000000000#',worker)
        self.assertNotIn('RPX_AccuracySolve',worker)
        self.assertIn('If original.mode <> "upper" Then Exit Function',worker)
        self.assertIn('seconds > 10# Then accepted = False',worker)
        self.assertIn('accepted = RPX_AccuracyDualAudit(original, updated)',worker)
        self.assertIn('If accepted Then candidate = updated',worker)
        self.assertNotRegex(worker,r'\bguideReady\s*=|\bxx\s*=|\bRStrengthFactor\s*=')

    def test_free_y_only_cone_bits_preserved_and_source_transaction(self):
        worker=procedure(self.new,'TryOriginalDualLinearRepair')
        self.assertIn('sourceDual = candidate',worker)
        self.assertIn('updated = sourceDual\n    For i = 0 To original.equalityCount - 1',worker)
        self.assertEqual(worker.count('candidate = updated'),1)
        self.assertIn('DualConeBitsUnchanged(sourceDual, updated, original.equalityCount, original.m)',worker)
        bits=procedure(self.new,'DualConeBitsUnchanged')
        self.assertIn('source(i) <> updated(i)',bits)
        self.assertIn('If source(i) = 0# Then',bits)
        self.assertIn('sourceBits <> updatedBits',bits)

    def test_transpose_uses_only_original_equality_rows_sorted(self):
        fn=procedure(self.new,'OriginalEqualityTranspose')
        self.assertIn('transposed.n = original.equalityCount: transposed.m = original.n',fn)
        self.assertIn('transposed.nnz = original.rowPtr(original.equalityCount)',fn)
        self.assertIn('For i = 0 To original.equalityCount - 1',fn)
        self.assertIn('transposed.colId(p) = i',fn)
        self.assertNotRegex(fn,r'\boriginal\.\w+\s*=')
        worker=procedure(self.new,'TryOriginalDualLinearRepair')
        self.assertIn('rhs(i) = -original.originalCost(i)',worker)
        self.assertIn('-original.values(j) * sourceDual(i)',worker)

    def test_flag0_and_auxiliary_regularization_scope(self):
        enabled=procedure(self.new,'DualLinearRepairEnabled')
        self.assertIn('RPX_Setting("ACCURACY_DLL_DUAL_LINEAR_REPAIR", 0)',enabled)
        reg=procedure(self.new,'SetUpperDualRegularization')
        self.assertIn('RPX_Setting("ACCURACY_DLL_DUAL_REGULARIZATION", 0)',reg)
        self.assertIn('dp.pdiag(i) = coefficient',reg)
        self.assertNotIn('originalCost',reg)
        helper=procedure(self.new,'RPX_AccuracyRecoverDual')
        self.assertIn('If mode = "upper" Then useDualRepair = DualLinearRepairEnabled()',helper)
        self.assertIn('If useDualRepair Then SetUpperDualRegularization dp',helper)
        self.assertIn('Erase helperDual',helper)

    def test_input_cancel_and_direct_exception_propagation(self):
        guard=procedure(self.new,'DualRepairAssert')
        for token in ['XCancel','err.Raise 18','RPX_AssertInputs stage','CurrentProgram(original)',
                      'P4Solve <> solveId','RPX_HSDCandidateId <> candidateId']:self.assertIn(token,guard)
        worker=procedure(self.new,'TryOriginalDualLinearRepair')
        self.assertIn('dual_repair_start_lsmr',worker);self.assertIn('dual_repair_after_original_audit',worker)
        direct=procedure(self.new,'RPX_AccuracyDirectBound')
        self.assertIn('If useDualRepairDirect And number <> 0 Then err.Raise',direct)
        self.assertLess(direct.index('TryOriginalDualLinearRepair'),direct.index('UpperRecoveryFallback:'))

    def test_diagnostic_source_capture_not_accepted_proof(self):
        capture=procedure(self.new,'CaptureDualRepairStage')
        self.assertIn('".dual_repair_"',capture)
        self.assertNotIn('".accuracy_',capture)
        self.assertIn('candidate_only=1',capture)
        worker=procedure(self.new,'TryOriginalDualLinearRepair')
        self.assertIn('sourceDual, "before", False',worker)
        self.assertIn('updated, "after", accepted',worker)

    def test_manifest_encoding_and_profile(self):
        m=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
        self.assertEqual(overlay.sha(overlay.BASE),m['baseline_sha256'])
        self.assertEqual(overlay.sha(ROOT/m['source']),m['source_sha256'])
        self.assertEqual(self.new,(ROOT/m['imported']).read_text(encoding='cp932'))
        self.assertEqual(m['profile']['ACCURACY_DLL_DUAL_LINEAR_REPAIR']['default'],0)
        self.assertEqual(m['profile']['ACCURACY_DLL_DUAL_SECONDS']['G1R20_precision_profile'],60)
        for name,h in m['files'].items():self.assertEqual(overlay.sha(ROOT/name),h)


if __name__=='__main__':unittest.main()
