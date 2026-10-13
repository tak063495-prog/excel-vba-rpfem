"""Strict original Farkas gates and unchanged primal evidence, repo-only fixtures."""
from pathlib import Path
import sys
from types import SimpleNamespace
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import original_farkas as farkas
from experiment_core import np,sparse


def specimen(h=-2.,mode='upper',q=8,cost=0.):
    # Analytic alternative: x=1, x+h>=0. h=-2 impossible, h=0 possible.
    m=SimpleNamespace(mesh_id='mesh',physical_model_id='model',provenance={'case':'analytic'})
    return SimpleNamespace(mode=mode,q=q,fs=1.,m=m,physical_x=np.array([1.]),
        original=(sparse.csc_matrix([[1.]]),np.array([1.]),sparse.csc_matrix([[1.]]),
                  np.array([h]),[1],np.array([cost])))


class OriginalFarkasTests(unittest.TestCase):
    def test_lower_nonzero_cost_and_non_q8_are_rejected(self):
        for pr in [specimen(mode='lower'),specimen(q=4),specimen(cost=1e-15)]:
            with self.subTest(pr=pr):
                result=farkas.solve_original_farkas(pr,seconds=1.)
                self.assertFalse(result['accepted']);self.assertNotIn('positive_sign',result)

    def test_nonfinite_or_invalid_raw_duals_never_supply_a_sign(self):
        for dual in [np.array([np.nan,1.]),np.array([1.,np.inf]),np.array([1.]),
                     np.array([1.,1.01]),np.array([-1.,-1.]),np.array([.5,.5]),np.ones((2,1))]:
            with self.subTest(dual=dual):
                result=farkas.audit_original_farkas(specimen(),dual)
                self.assertFalse(result['accepted']);self.assertNotIn('positive_sign',result)

    def test_tiny_direct_farkas_passes_all_original_gates_without_primal_changes(self):
        pr=specimen();before=[a.copy() if hasattr(a,'copy') else list(a) for a in pr.original]
        x=pr.physical_x.copy();result=farkas.solve_original_farkas(pr,seconds=5.)
        self.assertTrue(result['accepted']);self.assertEqual(result['positive_sign'],1)
        self.assertLessEqual(result['dual_stationarity'],1e-7)
        self.assertLessEqual(result['dual_cone_error'],1e-8)
        self.assertGreaterEqual(result['dual_objective'],1.000001)
        self.assertFalse(result['publishes_primal_bound']);self.assertNotIn('dual',result)
        np.testing.assert_array_equal(pr.physical_x,x)
        for old,new in zip(before,pr.original):
            if sparse.issparse(old):self.assertEqual((old!=new).nnz,0)
            else:np.testing.assert_array_equal(old,new)

    def test_feasible_primal_cannot_supply_a_positive_sign(self):
        result=farkas.solve_original_farkas(specimen(h=0.),seconds=5.)
        self.assertFalse(result['accepted']);self.assertNotIn('positive_sign',result)

    def test_nonfinite_original_program_is_rejected_before_solve(self):
        pr=specimen();pr.original[0].data[0]=np.nan
        result=farkas.solve_original_farkas(pr,seconds=1.)
        self.assertFalse(result['accepted']);self.assertEqual(result['reason'],'NONFINITE_ORIGINAL_PROGRAM')


if __name__=='__main__':unittest.main()
