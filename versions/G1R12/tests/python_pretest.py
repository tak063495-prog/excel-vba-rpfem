from pathlib import Path
import json, re, unittest
ROOT = Path(__file__).resolve().parents[1]
TOKENS = ['CONIC_STEP_STALLED', 'CONIC_NOT_CONVERGED', 'NEWTON_FACTORIZATION_FAILED', 'NT_NONINTERIOR_POINT', 'NT_CORRECTOR_SINGULAR']
def eligible(number, message, cancel=False, audit=False):
    if cancel or audit or number != 5: return False
    token = message.strip(' ')
    if token in TOKENS: return True
    m = re.fullmatch(r'NEWTON_FACTORIZATION_FAILED pivot=(-?[0-9]{1,10})', token)
    return bool(m and -(2**31) <= int(m[1]) < 2**31)
CASES = [(5,t,True) for t in TOKENS]
CASES += [(5,'NEWTON_FACTORIZATION_FAILED pivot='+p,True) for p in ['-1','0','123','2147483647','-2147483648']]
CASES += [(5,t,False) for t in ['NEWTON_FACTORIZATION_FAILED pivot=', 'NEWTON_FACTORIZATION_FAILED pivot=2147483648','NEWTON_FACTORIZATION_FAILED pivot=-2147483649','NEWTON_FACTORIZATION_FAILED pivot=foo','NEWTON_FACTORIZATION_FAILED pivot=123 extra','NEWTON_FACTORIZATION_FAILED_SUFFIX pivot=123','prefix NEWTON_FACTORIZATION_FAILED pivot=123','NEWTON_FACTORIZATION_FAILED pivot=+123','NEWTON_FACTORIZATION_FAILED pivot=1.5','NEWTON_FACTORIZATION_FAILED pivot=1\n','UNKNOWN','RAW_YIELD_AUDIT_FAILED','NT_NONINTERIOR_POINT extra']]
CASES += [(n,'NEWTON_FACTORIZATION_FAILED pivot=123',False) for n in [7,9,18,6,1004]]
class Pretests(unittest.TestCase):
    def test_token_grammar(self):
        for n,t,want in CASES: self.assertEqual(eligible(n,t),want,(n,t))
        for cancel,audit in [(True,False),(False,True)]: self.assertFalse(eligible(5,CASES[0][1],cancel,audit))
    def test_owner_counterexamples(self):
        a={'mesh':(601,1.),'field_owner':(601,1.),'lower':.60,'upper':.71}
        for bmesh in [(1000,2.),(601,2.)]:
            b={'mesh':bmesh,'lower':.65,'upper':.72}
            upper=min([a,b],key=lambda s:s['upper']); pair=min([a,b],key=lambda s:(s['upper']-s['lower'])/(s['upper']+s['lower']))
            self.assertNotEqual(upper['mesh'],pair['mesh'])
            output={'mesh':upper['mesh'],'field_owner':upper['field_owner']}
            self.assertEqual(output['mesh'],output['field_owner'])
    def test_actual_source_order_and_invariants(self):
        control=(ROOT/'src/RPX_Control.bas').read_text(encoding='utf-8')
        branch=control.split('        RPX_RunAdapt mode',1)[1].split('    Else\n        RAdaptRoute',1)[0]
        self.assertLess(branch.index('RPX_PrepareUpperOutput(True)'),branch.index('RPX_WriteMesh'))
        self.assertLess(branch.index('RPX_PrepareUpperOutput(True)'),branch.index('RPX_BearingMeasure'))
        changed=[]
        for p in (ROOT/'src').glob('*.bas'):
            if p.read_bytes() != (ROOT/'original'/p.name).read_bytes(): changed.append(p.name)
        self.assertEqual(sorted(changed),['RPX_Control.bas','RPX_Output.bas','RPX_Robust.bas'])
if __name__ == '__main__':
    (ROOT/'tests/token_cases.json').write_text(json.dumps(CASES,ensure_ascii=False,indent=2),encoding='utf-8')
    result=unittest.TextTestRunner(verbosity=2).run(unittest.defaultTestLoader.loadTestsFromTestCase(Pretests))
    (ROOT/'results/python_pretest.json').write_text(json.dumps({'scope':'Python simulation + source assertions; not Excel/physical acceptance','tests':result.testsRun,'failures':len(result.failures),'errors':len(result.errors),'grammar_vectors':len(CASES),'pass':result.wasSuccessful()},indent=2),encoding='utf-8')
    assert result.wasSuccessful()
