"""Analytic original-dual inequalities for upper and free-load lower programs."""
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import original_dual as dual
from experiment_core import np,sparse


def upper(bound):
    # min x subject to x>=bound: z=1 has dual objective bound.
    m=SimpleNamespace(mesh_id='mesh',physical_model_id='model',provenance={'case':'analytic'},tri=[[0,1,2]])
    return SimpleNamespace(mode='upper',q=8,fs=1.,m=m,physical_x=np.array([bound]),
        original=(sparse.csc_matrix((0,1)),np.zeros(0),sparse.csc_matrix([[1.]]),
                  np.array([-bound]),[1],np.array([1.])))


def lower(cap):
    # min -lambda subject to lambda<=cap, with the required 18ne+1 layout.
    m=upper(0.).m;G=sparse.csc_matrix(([-1.],([0],[18])),shape=(1,19));cost=np.zeros(19);cost[-1]=-1.
    return SimpleNamespace(mode='lower',q=8,fs=1.,m=m,physical_x=np.zeros(19),
        original=(sparse.csc_matrix((0,19)),np.zeros(0),G,np.array([cap]),[1],cost))


class OriginalDualTests(unittest.TestCase):
    def test_upper_minimum_above_one_gives_only_positive_discrete_sign(self):
        pr=upper(2.);x=pr.physical_x.copy();e=dual.solve_original_dual(pr,seconds=5.)
        self.assertTrue(e['accepted']);self.assertEqual(e['search_sign'],1)
        self.assertFalse(e['publishes_primal_bound']);np.testing.assert_array_equal(pr.physical_x,x)

    def test_upper_minimum_below_one_refuses_positive_sign(self):
        e=dual.solve_original_dual(upper(.5),seconds=5.)
        self.assertFalse(e['accepted']);self.assertNotIn('search_sign',e)

    def test_lower_maximum_below_one_gives_only_negative_discrete_sign(self):
        pr=lower(.5);x=pr.physical_x.copy();e=dual.solve_original_dual(pr,seconds=5.)
        self.assertTrue(e['accepted']);self.assertEqual(e['search_sign'],-1)
        self.assertGreaterEqual(e['dual_objective'],e['threshold']);np.testing.assert_array_equal(pr.physical_x,x)

    def test_lower_maximum_above_one_refuses_negative_sign(self):
        e=dual.solve_original_dual(lower(1.5),seconds=5.)
        self.assertFalse(e['accepted']);self.assertNotIn('search_sign',e)

    def test_lower_fixed_load_and_inexact_cost_or_layout_are_prohibited(self):
        cases=[]
        for value in [0.,-2.,-1.+1e-15]:
            pr=lower(.5);pr.original[-1][-1]=value;cases.append(pr)
        pr=upper(2.);pr.mode='lower';cases.append(pr)
        for pr in cases:
            with self.subTest(pr=pr):
                e=dual.solve_original_dual(pr,seconds=5.)
                self.assertFalse(e['accepted']);self.assertNotIn('search_sign',e)

    def test_nonfinite_stat_cone_and_insufficient_objective_are_rejected(self):
        for pr,z in [(upper(2.),np.array([np.nan])),(upper(2.),np.array([np.inf])),
                     (upper(2.),np.array([1.01])),(upper(2.),np.array([-1.])),
                     (upper(.5),np.array([1.])),(lower(1.5),np.array([1.]))]:
            with self.subTest(pr=pr,z=z):
                e=dual.audit_original_dual(pr,z)
                self.assertFalse(e['accepted']);self.assertNotIn('search_sign',e)


if __name__=='__main__':unittest.main()
