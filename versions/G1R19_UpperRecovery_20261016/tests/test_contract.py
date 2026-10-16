"""Source preservation and independent positive-work identity checks; solver free."""
from pathlib import Path
import importlib.util, json, math, re, sys, unittest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(Path(__file__).resolve().parent))
from normalization_reference import fixture, normalize

spec = importlib.util.spec_from_file_location('upper_overlay', ROOT / 'build_overlay.py')
overlay = importlib.util.module_from_spec(spec)
spec.loader.exec_module(overlay)


def procedure(text, name):
    match = re.search(r'^(?:Public|Private) (Function|Sub) ' + re.escape(name) + r'\b', text, re.M)
    if not match:
        raise ValueError(name)
    end = text.index('\nEnd ' + match[1], match.start()) + len('\nEnd ' + match[1])
    return text[match.start():end]


class Contracts(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.old = overlay.BASE.read_text(encoding='utf8')
        cls.new = (ROOT / 'src/vba/RPX_AccuracyNative.bas').read_text(encoding='utf8')

    def test_single_module_exact_controlled_patch(self):
        self.assertEqual([p.name for p in (ROOT / 'src/vba').glob('*.bas')], ['RPX_AccuracyNative.bas'])
        self.assertEqual(self.new, overlay.expected())
        restored = self.new.removesuffix(overlay.HELPERS).replace(overlay.BRANCH, '', 1)
        for before, after in reversed(overlay.direct_changes()):
            self.assertEqual(restored.count(after), 1)
            restored = restored.replace(after, before, 1)
        self.assertEqual(self.old, restored)

    def test_legacy_flag0_and_lower_body_exact(self):
        old = procedure(self.old, 'RPX_AccuracyRecoverPrimal')
        new = procedure(self.new, 'RPX_AccuracyRecoverPrimal')
        self.assertEqual(old, new.replace(overlay.BRANCH, '', 1))
        self.assertIn('If mode = "upper" Then\n        If UpperRecenterEnabled() Then', new)
        self.assertIn('RPX_Setting("ACCURACY_DLL_UPPER_RECENTER", 0)', self.new)

    def test_original_gates_and_commits_unchanged(self):
        names = ['CurrentProgram', 'RPX_AccuracyExport', 'RPX_AccuracyPrimalAudit',
                 'RPX_AccuracyDualAudit', 'RPX_AccuracyDualProblem', 'RPX_AccuracyRecoverDual',
                 'RPX_AccuracySolve', 'HomogeneousLowerWitness',
                 'CaptureAccuracyProof', 'ConsumeUpperDualFirst']
        for name in names:
            with self.subTest(name=name):
                self.assertEqual(procedure(self.old, name), procedure(self.new, name))

    def test_direct_restore_commit_and_positive_dual_order_preserved(self):
        old = procedure(self.old, 'RPX_AccuracyDirectBound')
        new = procedure(self.new, 'RPX_AccuracyDirectBound')
        # All existing code, including Restore and original_primal commit, is
        # exactly recoverable by deleting only the declared opt-in patch.
        for before, after in reversed(overlay.direct_changes()):
            new = new.replace(after, before, 1)
        self.assertEqual(old, new)
        direct = procedure(self.new, 'RPX_AccuracyDirectBound')
        self.assertLess(direct.index('RPX_AccuracyDualAudit'), direct.index('UpperRecoveryFallback:'))
        self.assertIn('If Not accepted And useUpperRecenter Then', direct)
        self.assertIn('RecoverUpperRecenterProgram(original, candidate, hasOriginalCenter, recentered)', direct)
        self.assertIn('If recenterInvoked And number <> 0 Then err.Raise', direct)
        self.assertIn('direct_upper_recenter_commit', direct)

    def test_two_qp_fresh_lp_and_modified_dual_never_a_guide(self):
        recipe = procedure(self.new, 'RecoverUpperRecenterProgram')
        self.assertIn('For trialIndex = 1 To 2', recipe)
        self.assertEqual(recipe.count('RPX_AccuracySolve('), 2)  # One call inside bounded loop.
        self.assertLess(recipe.index('RPX_AccuracySolve(original'), recipe.index('For trialIndex'))
        self.assertIn('center = candidate', recipe)
        self.assertIn('qp = original: AddBound qp, -1, 0.999998', recipe)
        self.assertNotIn('ScaleSolverCones', recipe)
        self.assertNotIn('RPX_AccuracyDualAudit', recipe)
        self.assertNotIn('guideReady', recipe)
        self.assertEqual(recipe.count('Erase modifiedDual'), 3)
        self.assertNotRegex(recipe, r'\bxx\s*=|\bRStrengthFactor\s*=|\bXStatus\s*=')
        self.assertLess(recipe.index('AuditUpperRecenter'), recipe.index('recovered = candidate'))

    def test_work_context_and_full_q8_gate_required(self):
        scope = procedure(self.new, 'UpperWorkRow')
        for term in ['original.mode <> "upper"', 'original.q <> 8', 'count <> 1',
                     'original.equalityCount To original.m - 1', 'original.rhs(i) <> 0#']:
            self.assertIn(term, scope)
        audit = procedure(self.new, 'AuditUpperRecenter')
        self.assertLess(audit.index('NormalizeUpperWork'), audit.index('RPX_AccuracyPrimalAudit'))
        self.assertLess(audit.index('RPX_AccuracyPrimalAudit'), audit.index('RPX_FieldGate'))
        self.assertIn('RPX_FieldGate("upper", candidate, original.factor, 8', audit)
        self.assertIn('AccuracyNativeValue > 0.999999', audit)
        self.assertEqual(audit.count('UpperRecenterAssert'), 2)

    def test_cancel_input_model_epoch_and_exception_restore(self):
        guard = procedure(self.new, 'UpperRecenterAssert')
        for term in ['XCancel', 'err.Raise 18', 'RPX_AssertInputs stage', 'CurrentProgram(original)',
                     'P4Solve <> solveId', 'RPX_HSDCandidateId <> candidateId']:
            self.assertIn(term, guard)
        recipe = procedure(self.new, 'RecoverUpperRecenterProgram')
        for call in ['upper_recenter_lp_start', 'upper_recenter_lp_after',
                     'upper_recenter_qp_start', 'upper_recenter_qp_after', 'upper_recenter_return']:
            self.assertIn(call, recipe)
        self.assertIn('RPX_DiagAuditRestore auditState: RPX_AuditRejected = oldRejected', recipe)
        self.assertIn('If number <> 0 Then err.Raise number, errorSource, message', recipe)
        self.assertIn('Resume Restore', recipe)

    def test_work_identity_for_positive_homogeneous_rescaling(self):
        p = fixture()
        y, ratio, accepted = normalize(p, [2., 2.])
        self.assertTrue(accepted)
        self.assertEqual(ratio, 2.)
        self.assertEqual(y, [1., 1.])
        self.assertEqual(sum(a*b for a,b in zip(p.rows[0], y)), p.rhs[0])
        self.assertEqual(sum(a*b for a,b in zip(p.rows[1], y)), 0.)
        self.assertGreaterEqual(-sum(a*b for a,b in zip(p.rows[2], y)), 0.)
        # Positive scaling does not repair an invalid homogeneous equality/cone.
        z, _, accepted = normalize(p, [3., 1.])
        self.assertTrue(accepted)
        self.assertNotEqual(sum(a*b for a,b in zip(p.rows[1], z)), 0.)

    def test_inhomogeneous_and_nonpositive_work_rejected(self):
        for index, value in [(1, .1), (2, 1e-300), (0, 0.)]:
            p = fixture(); p.rhs[index] = value
            self.assertFalse(normalize(p, [2., 2.])[2])
        for x in [[0.,0.], [-1.,-1.], [math.nan,1.], [math.inf,1.]]:
            self.assertFalse(normalize(fixture(), x)[2])
        for attr, value in [('mode', 'lower'), ('q', 4), ('phase', 'fixed_load')]:
            p = fixture(); setattr(p, attr, value)
            self.assertFalse(normalize(p, [2.,2.])[2])

    def test_manifest_and_encoding(self):
        manifest = json.loads((ROOT / 'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
        self.assertEqual(manifest['profile']['ACCURACY_DLL_UPPER_RECENTER']['default'], 0)
        for key in ['source', 'imported']:
            hash_key = 'source_sha256' if key == 'source' else 'import_sha256'
            self.assertEqual(overlay.sha(ROOT / manifest[key]), manifest[hash_key])
        self.assertEqual(self.new, (ROOT / manifest['imported']).read_text(encoding='cp932'))
        self.assertEqual(overlay.sha(overlay.BASE), manifest['baseline_sha256'])
        for path, expected in manifest['files'].items():
            self.assertEqual(overlay.sha(ROOT / path), expected)


if __name__ == '__main__':
    unittest.main()
