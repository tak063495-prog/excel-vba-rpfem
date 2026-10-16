"""Meaningful source contracts and signed-interval scheduling invariants."""
from pathlib import Path
import importlib.util,json,re,unittest
ROOT=Path(__file__).resolve().parents[1]
spec=importlib.util.spec_from_file_location('terminal_overlay',ROOT/'build_overlay.py')
overlay=importlib.util.module_from_spec(spec);spec.loader.exec_module(overlay)

def procedure(text,name):
    match=re.search(r'^(?:Public|Private) (Function|Sub) '+re.escape(name)+r'\b',text,re.M)
    if not match:raise ValueError(name)
    last=text.index('\nEnd '+match[1],match.start())+len('\nEnd '+match[1])
    return text[match.start():last]

class ContractTests(unittest.TestCase):
    def test_only_controlled_two_modules(self):
        self.assertEqual({p.name for p in (ROOT/'src/vba').glob('*.bas')},set(overlay.expected()))
        for name,expected in overlay.expected().items():self.assertEqual((ROOT/'src/vba'/name).read_text(encoding='utf8'),expected)

    def test_original_physical_and_root_gates_unchanged(self):
        pairs={'RPX_Fs.bas':['TryFsTrialPoint','RPX_ActualRootBracket','SaveBracket','BracketUsable','AccuracyFsRecover'],
               'RPX_AccuracyNative.bas':['CurrentProgram','RPX_AccuracyExport','RPX_AccuracyPrimalAudit','RPX_AccuracyDualAudit','RPX_AccuracyDualProblem','RPX_AccuracyRecoverDual','HomogeneousLowerWitness']}
        for name,functions in pairs.items():
            old=(overlay.BASE/name).read_text(encoding='utf8');new=(ROOT/'src/vba'/name).read_text(encoding='utf8')
            for function in functions:
                with self.subTest(function=function):self.assertEqual(procedure(old,function),procedure(new,function))

    def test_shared_original_commit_and_restore_unchanged(self):
        old=(overlay.BASE/'RPX_AccuracyNative.bas').read_text(encoding='utf8')
        new=(ROOT/'src/vba/RPX_AccuracyNative.bas').read_text(encoding='utf8')
        before=procedure(old,'RPX_AccuracyDirectBound');after=procedure(new,'RPX_AccuracyDirectBound')
        self.assertEqual(before[before.index('\nRestore:'):],after[after.index('\nRestore:'):])

    def test_final_endpoint_audit_and_root_completion_unchanged(self):
        old=procedure((overlay.BASE/'RPX_Fs.bas').read_text(encoding='utf8'),'RPX_FsEndpoint')
        new=procedure((ROOT/'src/vba/RPX_Fs.bas').read_text(encoding='utf8'),'RPX_FsEndpoint')
        anchor='    If i > 40 Then RFsSearchIncomplete = True:'
        self.assertEqual(old[old.index(anchor):],new[new.index(anchor):])

    def test_no_unbounded_retry_or_tolerance_change(self):
        text=(ROOT/'src/vba/RPX_Fs.bas').read_text(encoding='utf8')
        self.assertEqual(text.count('For i = 1 To 40'),1)
        self.assertIn('width <= 2.5 * RFsTolerance * valueScale',text)
        part=text[text.index('        If terminalActive And mode = "lower" Then'):text.index('        probe = False',text.index('        If terminalActive And mode = "lower" Then'))]
        self.assertEqual(part.count('TryFsTrialPoint('),2)
        self.assertNotIn('Do While',part)
        self.assertIn('If trialA.valid Then',part);self.assertIn('If trialB.valid Then',part)
        self.assertIn('FS_TERMINAL_BOTH_UNKNOWN',part)

    def test_request_context_and_all_exit_resets(self):
        native=(ROOT/'src/vba/RPX_AccuracyNative.bas').read_text(encoding='utf8')
        consume=procedure(native,'ConsumeUpperDualFirst')
        for term in ['upperFirstFactor','upperFirstEpoch','upperFirstQ','upperFirstModel','upperFirstPolicy','upperFirstPhase','upperFirstSolve','upperFirstCandidate']:
            self.assertIn(term,consume)
        self.assertIn('Call RPX_AccuracyUpperDualFirstReset',consume)
        fs=(ROOT/'src/vba/RPX_Fs.bas').read_text(encoding='utf8')
        self.assertEqual(procedure(fs,'TryFsValue').count('Call RPX_AccuracyUpperDualFirstReset'),3)
        self.assertIn('Call RPX_AccuracyUpperDualFirstReset',procedure(fs,'RPX_FsBracketReset'))

    def test_manifest_import_and_profile_defaults(self):
        m=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
        self.assertEqual(m['profile']['FS_TERMINAL_SCHEDULE']['default'],0)
        self.assertEqual(m['profile']['G1R20_precision_profile']['DIRECT_DLL_SECONDS'],30)
        self.assertEqual(m['profile']['G1R20_precision_profile']['ACCURACY_DLL_DUAL_SECONDS'],30)
        for module in m['modules']:
            src=ROOT/module['source'];imp=ROOT/module['imported']
            self.assertEqual(overlay.sha(src),module['source_sha256']);self.assertEqual(overlay.sha(imp),module['import_sha256'])
            self.assertEqual(src.read_text(encoding='utf8'),imp.read_text(encoding='cp932'))
        for name,expected in m['files'].items():self.assertEqual(overlay.sha(ROOT/name),expected)

    def test_two_fraction_interval_monotonicity_and_unknown(self):
        # Exhaust all accepted sign/UNKNOWN combinations. The scheduler must
        # preserve bracketing and never evaluate the second point outside it.
        for first in [None,1,-1]:
            for second in [None,1,-1]:
                a,b=1.,1.00024;source_a,source_b=a,b;visited=[]
                for fraction,sign in [(.3,first),(.7,second)]:
                    x=source_a+fraction*(source_b-source_a)
                    if not a<x<b:continue
                    visited.append(x)
                    if sign==1:a=x
                    elif sign==-1:b=x
                self.assertGreater(b,a);self.assertGreaterEqual(a,source_a);self.assertLessEqual(b,source_b)
                if first is None and second is None:self.assertEqual((a,b),(source_a,source_b))
                if first==-1:self.assertEqual(len(visited),1)

if __name__=='__main__':unittest.main()
