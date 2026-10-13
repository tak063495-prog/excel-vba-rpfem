"""Local polynomial reproduction and a frozen P06N original-gate regression."""
from pathlib import Path
import hashlib
import json
import sys
import unittest
from collections import Counter

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import accuracy_driver as driver
import experiment_core as core
from mesh_geometry import interpolate_velocity
from precision_interpolation import local_barycentric,sample_p2,precision_prolong_upper
from model_support.schema import Mesh
from experiment_core import np

FIXTURE=Path(__file__).resolve().parents[1]/'python/data/frozen_upper_interpolation'


def mesh(xy,tri,material_ids=None):
    ne=len(tri)
    counts=Counter(tuple(sorted((int(a),int(b)))) for t in tri for a,b in zip(t,np.roll(t,-1)))
    boundary={edge:{'kind':'load','rigid':-1,'loads':np.zeros((2,2))} for edge,count in counts.items() if count==1}
    return Mesh(np.asarray(xy,float),np.asarray(tri,np.int32),
                np.array([[1,30,30,0],[2,30,30,0]],float),
                np.zeros(ne,np.int32) if material_ids is None else np.asarray(material_ids,np.int32),
                np.full(ne,-1),np.zeros((ne,2,2)),boundary,1.,'synthetic_same_model',provenance={}).validate()


def nodes(m):
    p=m.xy[m.tri]
    return np.concatenate([p,(p+np.roll(p,-1,axis=1))/2],axis=1)


class PrecisionInterpolationTests(unittest.TestCase):
    def nested_pair(self,shift=0.):
        points=np.array([[0,0],[2,0],[0,2],[1,0],[1,1],[0,1]],float)+shift
        parent=mesh(points[:3],[[0,1,2]])
        child=mesh(points,[[0,3,5],[3,1,4],[5,4,2],[3,4,5]])
        return parent,child

    def test_constant_velocity_is_copied_exactly_on_translated_mesh(self):
        parent,child=self.nested_pair(2.**24)
        values=np.tile(np.array([1e6,-3.125]),(6,1)).ravel()
        result=precision_prolong_upper(parent,child,values).reshape(-1,2)
        np.testing.assert_array_equal(result,np.tile([1e6,-3.125],(24,1)))

    def test_full_quadratic_polynomial_reproduction_and_same_parent_mapping(self):
        for shift in [0.,2.**24]:
            with self.subTest(shift=shift):
                parent,child=self.nested_pair(shift)
                def polynomial(points):
                    x,y=(points-shift).T
                    return np.column_stack([x*x+2*x*y-3*y*y+5*x+7*y+11,
                                            -2*x*x+3*x*y+y*y-4*x+5*y-13])
                source=polynomial(nodes(parent).reshape(-1,2)).ravel()
                result,ids=precision_prolong_upper(parent,child,source,return_parent_ids=True)
                np.testing.assert_allclose(result.reshape(-1,2),polynomial(nodes(child).reshape(-1,2)),rtol=0.,atol=2e-14)
                self.assertEqual(ids,[0]*4)

    def test_non_nested_same_material_crossing_and_changed_material_are_rejected(self):
        xy=[[0,0],[1,0],[1,1],[0,1]]
        parent=mesh(xy,[[0,1,2],[0,2,3]])
        crossing=mesh(xy,[[0,1,3],[1,2,3]])
        with self.assertRaisesRegex(ValueError,'not contained'):
            precision_prolong_upper(parent,crossing,np.zeros(24))
        p,c=self.nested_pair();c.material_id[:]=1
        with self.assertRaisesRegex(ValueError,'same-material'):
            precision_prolong_upper(p,c,np.zeros(12))

    def test_nonfinite_degenerate_and_wrong_model_candidates_are_rejected(self):
        with self.assertRaises(ValueError):local_barycentric([[0,0],[1,0],[2,0]],[1,0])
        with self.assertRaises(ValueError):sample_p2(np.full((6,2),np.nan),[1,0,0])
        p,c=self.nested_pair();c.physical_model_id='changed'
        with self.assertRaises(ValueError):precision_prolong_upper(p,c,np.zeros(12))

    def test_frozen_original_raw_gate_rejects_legacy_and_accepts_local_interpolation(self):
        manifest=json.loads((FIXTURE/'manifest.json').read_text(encoding='utf8'))
        for filename,expected in manifest['sha256'].items():
            self.assertEqual(hashlib.sha256((FIXTURE/filename).read_bytes()).hexdigest(),expected)
        record=json.loads((FIXTURE/'P06N.json').read_text(encoding='utf8'))
        parent=driver.load_saved_mesh(FIXTURE/'source_mesh.npz',record)
        child=driver.load_saved_mesh(FIXTURE/'child_mesh.npz',record)
        with np.load(FIXTURE/'source_field.npz',allow_pickle=False) as d:source=d['x'].copy();fs=float(d['fs'])
        with np.load(FIXTURE/'legacy_global_inverse_field.npz',allow_pickle=False) as d:legacy=d['x'].copy()
        prepared=core.Prepared(child,fs,'upper',8,reduction='none')
        original_before=[a.copy() if hasattr(a,'copy') else list(a) for a in prepared.original]
        old=core.evaluate(prepared,legacy)
        self.assertFalse(old['physical_pass']);self.assertGreater(old['cone_error'],1e-7)
        for label,candidate in [('nonnested',interpolate_velocity(parent,child,source)),
                                ('optional_nested',precision_prolong_upper(parent,child,source))]:
            with self.subTest(method=label):
                result=core.evaluate(prepared,candidate)
                self.assertTrue(result['physical_pass']);self.assertTrue(result['witness'])
                self.assertLess(result['cone_error'],1e-7)
                self.assertLess(abs(result['value']-manifest['source_value']),2e-12)
        for before,after in zip(original_before,prepared.original):
            if hasattr(before,'toarray'):
                self.assertEqual((before!=after).nnz,0)
            else:np.testing.assert_array_equal(before,after)


if __name__=='__main__':unittest.main()
