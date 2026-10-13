"""Analytic alternative-system checks; solver status never substitutes residuals."""
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import diagnose_farkas as diagnosis
from experiment_core import np,sparse


def problem(h=-2.,mode='upper',q=8,cost=0.):
    # x=1, x+h>=0. h=-2 is infeasible; h=0 is feasible.
    m=SimpleNamespace(provenance={'case':'analytic'},mesh_id='mesh',physical_model_id='model',tri=[[0,1,2]])
    return SimpleNamespace(mode=mode,q=q,fs=1.,m=m,
        original=(sparse.csc_matrix([[1.]]),np.array([1.]),sparse.csc_matrix([[1.]]),
                  np.array([h]),[1],np.array([cost])))


class FarkasDiagnosisTests(unittest.TestCase):
    def test_analytic_positive_witness_is_directly_reaudited(self):
        target=1.000003;result=diagnosis.audit_farkas(problem(),np.array([target,target]))
        self.assertTrue(result['accepted']);self.assertEqual(result['positive_discrete_sign'],1)
        self.assertFalse(result['primal_bound'])

    def test_wrong_mode_q_nonzero_cost_are_ineligible(self):
        for pr in [problem(mode='lower'),problem(q=4),problem(cost=1e-15)]:
            with self.assertRaises(ValueError):diagnosis.audit_farkas(pr,np.ones(2))

    def test_original_stationarity_cone_and_separation_each_reject(self):
        for w in [np.array([1.,1.001]),np.array([-1.,-1.]),np.array([.5,.5]),
                  np.array([np.nan,1.]),np.array([1.])]:
            with self.subTest(w=w):
                self.assertFalse(diagnosis.audit_farkas(problem(),w)['accepted'])

    def test_direct_feasibility_distinguishes_both_analytic_sides(self):
        infeasible,_=diagnosis.direct_farkas(problem(),time_limit=5.)
        feasible,_=diagnosis.direct_farkas(problem(h=0.),time_limit=5.)
        self.assertTrue(infeasible['accepted']);self.assertFalse(feasible['accepted'])


if __name__=='__main__':unittest.main()
