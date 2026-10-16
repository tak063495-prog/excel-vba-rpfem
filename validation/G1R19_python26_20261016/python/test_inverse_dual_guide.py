"""Real assembled-program duals for the inverse-only polish shortcut."""
import unittest
import experiment_core as core
from experiment_core import np
from original_dual import solve_original_dual, audit_original_dual


def physical_triangle():
    boundary={(0,1):{'kind':'fixed','rigid':-1,'loads':np.zeros((2,2))},
              (1,2):{'kind':'load','rigid':-1,'loads':np.zeros((2,2))},
              (0,2):{'kind':'load','rigid':-1,'loads':np.zeros((2,2))}}
    return core.Mesh(np.array([[0.,0.],[1.,0.],[0.,1.]]),np.array([[0,1,2]],np.int32),
        np.array([[1.,0.,0.,1.]]),np.array([0],np.int32),np.array([-1],np.int32),
        np.array([[[0.,0.],[0.,-1.]]]),boundary,1.,'real_triangle_program',
        provenance={'psi_policy':'hold_initial','case':'TEST_DUAL_GUIDE'}).validate()


class InverseDualGuideTests(unittest.TestCase):
    def test_actual_original_duals_supply_only_the_inverse_side(self):
        for mode,fs in [('lower',100.),('upper',.01)]:
            with self.subTest(mode=mode):
                prepared=core.Prepared(physical_triangle(),fs,mode,8)
                self.assertFalse(prepared.physical_original_objective.flags.writeable)
                evidence=solve_original_dual(prepared,seconds=10,return_dual=True)
                self.assertTrue(evidence['accepted'])
                original=audit_original_dual(prepared,evidence['dual'])
                guide=core.inverse_dual_guide(prepared,original)
                self.assertTrue(guide['verified'])
                self.assertLess(guide['dual_stationarity'],1e-12)
                self.assertEqual(guide['dual_cone_error'],0)

    def test_missing_nonfinite_and_zero_duals_cannot_skip_recovery(self):
        for mode,fs in [('lower',100.),('upper',.01)]:
            prepared=core.Prepared(physical_triangle(),fs,mode,8)
            E,b,G,h,*_=prepared.original
            dual=np.zeros(len(b)+len(h))
            self.assertFalse(core.inverse_dual_guide(prepared,audit_original_dual(prepared,dual))['verified'])
            dual[0]=np.nan
            self.assertFalse(core.inverse_dual_guide(prepared,audit_original_dual(prepared,dual))['verified'])
            self.assertFalse(core.inverse_dual_guide(prepared,{})['verified'])

    def test_fixed_load_and_changed_upper_objectives_reject_replayed_guides(self):
        mesh=physical_triangle()
        original=core.Prepared(mesh,100.,'lower',8)
        evidence=solve_original_dual(original,seconds=10)
        fixed=core.Prepared(mesh,100.,'lower',8,fixed_load=1.000003)
        guide=core.inverse_dual_guide(fixed,evidence)
        self.assertFalse(guide['verified'])
        self.assertEqual(guide['reason'],'LOWER_FREE_LOAD_OBJECTIVE_AND_LAYOUT_REQUIRED')
        upper=core.Prepared(mesh,.01,'upper',8)
        evidence=solve_original_dual(upper,seconds=10)
        E,b,G,h,dims,cost=upper.original
        upper.original=E,b,G,h,dims,np.zeros(len(cost))
        guide=core.inverse_dual_guide(upper,evidence)
        self.assertFalse(guide['verified'])
        self.assertEqual(guide['reason'],'ORIGINAL_UPPER_OBJECTIVE_AND_LAYOUT_REQUIRED')

    def test_genuine_primal_bounds_retain_the_recovery_path(self):
        for mode,fs in [('lower',.01),('upper',100.)]:
            with self.subTest(mode=mode):
                prepared=core.Prepared(physical_triangle(),fs,mode,8)
                result,field,_=core.solve(prepared,max_iter=350,primal_polish=True)
                self.assertTrue(result['witness'])
                self.assertTrue(core.evaluate(prepared,field)['physical_pass'])
                self.assertFalse(core.inverse_dual_guide(prepared,result)['verified'])


if __name__=='__main__':unittest.main()
