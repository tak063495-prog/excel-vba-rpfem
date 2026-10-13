"""Original-program Farkas ray recovery; synthetic models and temporary fields."""
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

PYTHON=Path(__file__).resolve().parents[1]/'python'
sys.path.insert(0,str(PYTHON))
import experiment_core as core
import accuracy_driver as driver
from experiment_core import np,sparse


def prepared(cost=0.,mode='upper',b=-1.,E=1.,G=1.):
    raw=(sparse.csc_matrix([[E]]),np.array([b]),sparse.csc_matrix([[G]]),
         np.zeros(1),[1],np.array([cost]))
    m=SimpleNamespace(tri=np.array([[0,1,2]]),provenance={'case':'ray-test'},stress_scale=1.)
    return SimpleNamespace(original=raw,reduced=raw,cost=raw[-1],
        A=sparse.vstack([raw[0],-raw[2]],format='csc'),rhs=np.array([b,0.]),
        cones=[core.clarabel.ZeroConeT(1),core.clarabel.NonnegativeConeT(1)],
        objective_offset=0.,assembly_seconds=0.,transform_seconds=0.,m=m,
        fs=1.,mode=mode,q=8,kind='none',restore=lambda x:np.asarray(x).copy(),
        restore_dual=lambda z:(z.copy(),{'seconds':0.}))


class ZeroCostRayRecoveryTests(unittest.TestCase):
    def solve_fake(self,pr,z=(1.,1.),status='PrimalInfeasible',x=(0.,),**kwargs):
        class Solver:
            def __init__(self,*args):pass
            def solve(self):
                return SimpleNamespace(x=np.array(x),z=np.array(z),status=status,
                    iterations=1,solve_time=0.,obj_val=0.,obj_val_dual=0.)
        with patch.object(core.clarabel,'DefaultSolver',Solver),patch.object(core,'audit',
                return_value={'pass':False,'value':0.,'errors':{}}):
            return core.solve(pr,**kwargs)

    def test_real_infeasible_zero_cost_cone_recovers_original_ray_only(self):
        with patch.object(core,'audit',return_value={'pass':False,'value':0.,'errors':{}}), \
                tempfile.TemporaryDirectory(prefix='rpfem-ray-test-') as temporary:
            path=Path(temporary)/'ray.npz'
            data,x,z=core.solve(prepared(),save=path)
            with np.load(path) as saved:
                np.testing.assert_array_equal(saved['x'],x)
                np.testing.assert_array_equal(saved['z'],z)
                np.testing.assert_array_equal(saved['zr'],z)
        self.assertIn(data['status'],('PrimalInfeasible','AlmostPrimalInfeasible'))
        recovery=data['original_dual_ray_rescaling']
        self.assertTrue(recovery['accepted'])
        self.assertAlmostEqual(recovery['factor'],2.)
        self.assertAlmostEqual(data['dual_objective'],2.,places=7)
        self.assertLessEqual(data['dual_stationarity'],1e-7)
        self.assertLessEqual(data['dual_cone_error'],1e-8)
        self.assertAlmostEqual(data['reduced_dual_objective'],data['dual_objective'])
        self.assertEqual(driver.sign(data,'upper'),1)
        self.assertFalse(data['physical_pass'])
        self.assertFalse(data['witness'])
        self.assertFalse(data['numerical_gate'])
        self.assertEqual(recovery['scope'],'DISCRETE_UPPER_INFEASIBILITY_GUIDE_ONLY')

    def test_other_cost_mode_status_or_duals_disabled_are_unchanged(self):
        cases=[(prepared(cost=1e-15),{},'PrimalInfeasible'),
               (prepared(mode='lower'),{},'PrimalInfeasible'),
               (prepared(),{'duals':False},'PrimalInfeasible')]
        cases += [(prepared(),{},status) for status in
                  ('Solved','AlmostSolved','DualInfeasible','AlmostDualInfeasible','NumericalError')]
        for pr,kwargs,status in cases:
            with self.subTest(mode=pr.mode,cost=pr.cost,status=status,kwargs=kwargs):
                data,x,z=self.solve_fake(pr,status=status,**kwargs)
                self.assertNotIn('original_dual_ray_rescaling',data)
                np.testing.assert_array_equal(z,[1.,1.])

    def test_candidate_stationarity_and_cone_error_must_pass_original_thresholds(self):
        tests=[(prepared(),[1.+.75e-7,1.],'stationarity'),
               (prepared(E=0.,G=0.),[1.,-.75e-8],'cone_error')]
        for pr,z,key in tests:
            with self.subTest(key=key):
                data,x,returned=self.solve_fake(pr,z=z)
                attempt=data['original_dual_ray_rescaling']
                self.assertFalse(attempt['accepted'])
                self.assertGreater(attempt['after'][key],1e-7 if key=='stationarity' else 1e-8)
                np.testing.assert_array_equal(returned,z)
                self.assertEqual(data['dual_objective'],attempt['before']['objective'])

    def test_inside_stationarity_and_cone_boundaries_are_accepted(self):
        for pr,z in [(prepared(),[1.+.49e-7,1.]),
                     (prepared(E=0.,G=0.),[1.,-.49e-8])]:
            with self.subTest(z=z):
                data,x,returned=self.solve_fake(pr,z=z,status='AlmostPrimalInfeasible')
                self.assertTrue(data['original_dual_ray_rescaling']['accepted'])
                np.testing.assert_array_equal(returned,2*np.array(z))

    def test_negative_or_zero_separation_is_not_an_improvement(self):
        for b in (1.,0.):
            data,x,z=self.solve_fake(prepared(b=b))
            self.assertFalse(data['original_dual_ray_rescaling']['accepted'])
            np.testing.assert_array_equal(z,[1.,1.])

    def test_overflow_or_nonfinite_original_ray_never_promotes_sign(self):
        data,x,z=self.solve_fake(prepared(),z=[1.1e308,1.1e308])
        self.assertFalse(data['original_dual_ray_rescaling']['accepted'])
        self.assertIsNone(data['original_dual_ray_rescaling']['after'])
        np.testing.assert_array_equal(z,[1.1e308,1.1e308])
        for bad in (np.nan,np.inf,-np.inf):
            data,x,z=self.solve_fake(prepared(),z=[bad,1.])
            self.assertNotIn('original_dual_ray_rescaling',data)
            self.assertIsNone(driver.sign(data,'upper'))

    def test_raw_dual_and_canonical_scaled_dual_remain_consistent(self):
        pr=prepared(G=3.)
        data,x,z=self.solve_fake(pr,z=[3.,3.],cone_scaling=True)
        self.assertTrue(data['original_dual_ray_rescaling']['accepted'])
        np.testing.assert_array_equal(z,[6.,2.])
        self.assertEqual(data['dual_stationarity'],0.)
        self.assertEqual(data['reduced_dual_stationarity'],0.)
        self.assertEqual(data['dual_objective'],6.)

    def test_nonfinite_primal_stays_unknown_despite_valid_ray_guide(self):
        data,x,z=self.solve_fake(prepared(),x=[np.nan])
        self.assertTrue(data['original_dual_ray_rescaling']['accepted'])
        self.assertTrue(np.isnan(x).all())
        self.assertFalse(data['physical_pass'])
        self.assertFalse(data['witness'])
        self.assertFalse(data['numerical_gate'])
        self.assertEqual(driver.sign(data,'upper'),1)


if __name__=='__main__':unittest.main(verbosity=2)
