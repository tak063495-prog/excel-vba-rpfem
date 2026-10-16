"""Static purity, serialization and independent spectral-function checks."""
from pathlib import Path
import json,re,unittest
import reference_fixtures as ref
np=ref.np
ROOT=Path(__file__).resolve().parents[1]

class ContractTests(unittest.TestCase):
    def test_pure_module_has_no_project_model_or_solver_dependency(self):
        text=(ROOT/'src/vba/RPX_FlowMesh.bas').read_text(encoding='utf8')
        code='\n'.join(line for line in text.splitlines() if not line.lstrip().startswith("'"))
        for token in ['RXY','RTri','RStrengthFactor','XStatus','RProgramKind','RSubdivisions','RPX_AssertInputs','RPX_Compile','RPX_Optimize','RPX_FieldGate']:
            self.assertIsNone(re.search(r'\b'+token+r'\b',code),token)
        self.assertEqual(re.findall(r'^Public (?:Function|Sub) (\w+)',text,re.M),['RPX_FlowMeshCandidate'])
        self.assertNotIn('Declare ',code)

    def test_exact_source_import_and_manifest(self):
        manifest=json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
        src=ROOT/manifest['module']['source'];imp=ROOT/manifest['module']['imported']
        self.assertEqual(ref.sha(src),manifest['module']['source_sha256'])
        self.assertEqual(ref.sha(imp),manifest['module']['import_sha256'])
        self.assertEqual(src.read_text(encoding='utf8'),imp.read_text(encoding='cp932'))
        for name,expected in manifest['files'].items():self.assertEqual(ref.sha(ROOT/name),expected,name)

    def test_analytic_spectral_floor_matches_eigh(self):
        rng=np.random.default_rng(20261016)
        for jac in [np.zeros((2,2)),np.eye(2),np.diag([1e-14,1.]),*list(rng.normal(size=(1000,2,2)))]:
            matrix=jac.T@jac;a,b,c=matrix[0,0],matrix[0,1],matrix[1,1]
            largest=max(abs(a),abs(b),abs(c));disc=largest*np.sqrt(((a-c)/largest)**2+4*(b/largest)**2) if largest else 0.
            high=(a+c)/2+disc/2;low=(a+c)/2-disc/2;scale=max(high,1e-20);floor=scale/9
            if low>=floor:actual=matrix/scale
            elif high<=floor:actual=np.eye(2)/9
            else:
                alpha=(high-floor)/(high-low);beta=floor-alpha*low
                actual=(alpha*matrix+beta*np.eye(2))/scale
            values,vectors=np.linalg.eigh(matrix);hi=max(values[-1],1e-20)
            expected=(vectors*np.maximum(values,hi/9))@vectors.T/hi
            np.testing.assert_allclose(actual,expected,rtol=8e-15,atol=8e-16)

    def test_fixture_bytes_and_protection_expectations(self):
        root=ROOT/'tests/fixtures';manifest=json.loads((root/'FIXTURES.json').read_text())
        self.assertEqual(len(manifest['fixtures']),3)
        for entry in manifest['fixtures']:
            with self.subTest(name=entry['name']):
                self.assertEqual(ref.sha(root/entry['input']),entry['input_sha256'])
                xy,tri,mat,v,guards,passes=ref.read_input(root/entry['input'])
                self.assertEqual(len(v),12*len(tri));self.assertEqual(passes,5)
                with np.load(root/entry['expected_tri'],allow_pickle=False) as data:expected=data['tri']
                edges={tuple(sorted((int(a),int(b)))) for t in expected for a,b in zip(t,np.roll(t,-1))}
                self.assertTrue(set(map(tuple,guards)).issubset(edges))
                self.assertTrue(entry['source_physical_audit']['physical_pass'])
                self.assertGreater(entry['flips'],0)

if __name__=='__main__':unittest.main()
