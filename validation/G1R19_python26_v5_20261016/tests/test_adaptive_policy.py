"""Light interpolation/selection regressions; no large solve or bound claim."""
from pathlib import Path
import copy
import sys
import tempfile
import unittest
from unittest.mock import patch
import numpy as np

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import accuracy_driver as driver
import experiment_core as core
import adaptive_policy as policy
from test_accuracy_budget import args, mesh as control_mesh, result as control_result


def unit_mesh(phi=30.):
    rec={'case':'UNIT','source_case':{'materials':[dict(c=10.,phi=phi,psi=phi,gamma=20.)],
         'policy':'equal_reduced_phi'},'region_areas':[1.],
         'boundary_explicit':[dict(a=a,b=b,kind='自由',load0=[0.,0.],load1=[0.,0.])
            for a,b in [([0.,0.],[1.,0.]),([1.,0.],[1.,1.]),([1.,1.],[0.,1.]),([0.,1.],[0.,0.])]]}
    xy=np.array([[0.,0.],[1.,0.],[1.,1.],[0.,1.]])
    tri=np.array([[0,1,2],[0,2,3]],dtype=np.int32)
    return core.mesh_from_arrays(xy,tri,np.zeros(2,dtype=np.int32),rec),rec


class AdaptiveTests(unittest.TestCase):
    def test_lower_polynomial_and_source_preserved(self):
        parent,rec=unit_mesh();child=core.uniform_refine(parent,rec)
        rng=np.random.default_rng(24)
        source=np.r_[rng.normal(size=36),1.000003];before=source.copy()
        got=policy.stable_lower_prolong(parent,child,source)
        np.testing.assert_array_equal(source,before)
        self.assertEqual(got[-1],source[-1])
        sample=np.array([.2,.3,.5])
        for e,p in enumerate(child.provenance['parent_element_ids']):
            point=sample@child.xy[child.tri[e]]
            oldb=policy.local_barycentric(parent.xy[parent.tri[p]],point)
            def bernstein(b):
                a,b,c=b;return np.array([a*a,b*b,c*c,2*a*b,2*b*c,2*c*a])
            expected=bernstein(oldb)@source[:-1].reshape(-1,6,3)[p]
            actual=bernstein(sample)@got[:-1].reshape(-1,6,3)[e]
            np.testing.assert_allclose(actual,expected,atol=2e-15,rtol=2e-15)

    def test_upper_quadratic_samples_and_identity_guard(self):
        parent,rec=unit_mesh();child=core.uniform_refine(parent,rec)
        def values(points):
            x,y=points.T
            return np.column_stack((1+x+x*y+2*y*y,2-x*x+3*y))
        oldpoints=parent.xy[parent.tri]
        oldpoints=np.concatenate([oldpoints,(oldpoints+np.roll(oldpoints,-1,axis=1))/2],axis=1)
        source=values(oldpoints.reshape(-1,2)).ravel();before=source.copy()
        got=policy.stable_prolong(parent,child,source,'upper')
        points=child.xy[child.tri]
        points=np.concatenate([points,(points+np.roll(points,-1,axis=1))/2],axis=1)
        np.testing.assert_allclose(got,values(points.reshape(-1,2)).ravel(),atol=2e-15,rtol=2e-15)
        np.testing.assert_array_equal(source,before)
        bad=copy.deepcopy(child);bad.physical_model_id='different'
        with self.assertRaises(ValueError):policy.stable_prolong(parent,bad,source,'upper')
        with self.assertRaises(ValueError):policy.stable_lower_prolong(parent,bad,np.r_[np.zeros(36),1.])

    def test_dissipation_decomposition_matches_original_cost(self):
        parent,_=unit_mesh()
        rng=np.random.default_rng(18);velocity=rng.normal(size=24)
        score,meta=policy.upper_refinement_score(parent,velocity,1.2,.2)
        self.assertLess(meta['decomposition_error'],1e-12)
        self.assertTrue(np.isfinite(score).all())
        self.assertTrue(np.all(score>=0))
        self.assertFalse(meta['certificate'])
        with self.assertRaises(ValueError):policy.upper_refinement_score(parent,velocity,1.2,1.01)

    def test_phi0_uses_existing_indicator(self):
        parent,_=unit_mesh(0.)
        velocity=np.arange(24,dtype=float)
        score,meta=policy.upper_refinement_score(parent,velocity,1.2,.2)
        np.testing.assert_array_equal(score,core.indicator(parent,velocity))
        self.assertEqual(meta['method'],'existing_kinematic_phi0_fallback')

    def test_stops_after_lower_update_without_next_upper_solve(self):
        m=control_mesh();calls=[]
        upper=(2+.0201)/(2-.0201)
        def search(mesh,mode,initial,folder,label,*pos,**kwargs):
            calls.append(mode)
            fs=(1. if calls.count('lower')==1 else 1.001) if mode=='lower' else upper
            if calls.count('upper')>1:raise AssertionError('Unneeded upper solve after target')
            cert=driver.Certificate(mode,fs,m,np.ones(19 if mode=='lower' else 12),
                control_result(True,1.01 if mode=='lower' else .99),'control only',True)
            return cert,{'root_complete':True}
        with tempfile.TemporaryDirectory() as out,patch.object(driver,'load_model',return_value=(m,{})),\
                patch.object(driver,'search',side_effect=search),\
                patch.object(core,'local_refine',return_value=(m,{'marked':1})),\
                patch.object(core,'indicator',return_value=np.ones(1)),\
                patch.object(policy,'stable_prolong',return_value=np.ones(12)):
            report=driver.run(args(out,cycles=1,upper_indicator='kinematic'))
        self.assertEqual(calls,['lower','upper','lower'])
        self.assertTrue(report['target_met']);self.assertEqual(report['status'],'PASS')
        self.assertEqual(report['reason'],'target_met_after_lower')


if __name__=='__main__':unittest.main()
