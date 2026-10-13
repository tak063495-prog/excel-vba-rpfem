"""Analytic full-vector transport roots, with no scalar-dual substitution."""
from pathlib import Path
import sys
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import zero_phi_signed_transport as transport
import experiment_core as core
from check_zero_phi import specimen
from experiment_core import np


class SignedPhiZeroTests(unittest.TestCase):
    def source(self,mode):
        m=specimen();pr=core.Prepared(m,1.,mode,8,reduction='none')
        result,x,z=core.solve(pr,{'equilibrate_enable':True})
        self.assertTrue(result['physical_pass'])
        return m,x,z

    def test_both_analytic_sides_close_with_fresh_original_endpoint_gates(self):
        for mode in ['lower','upper']:
            with self.subTest(mode=mode):
                m,x,z=self.source(mode);xx=x.copy();zz=z.copy()
                evidence=transport.signed_transport(m,1.,x,z,mode)
                self.assertTrue(evidence['root']['root_complete'])
                self.assertEqual(evidence['root']['bracket_signs'],[1,-1])
                self.assertLess(evidence['root']['bracket_width'],1e-4)
                self.assertTrue(evidence['selected']['audit']['witness'])
                self.assertLess(abs(evidence['selected']['fs']-2.),2e-5)
                np.testing.assert_array_equal(x,xx);np.testing.assert_array_equal(z,zz)
                for p in evidence['endpoints']:
                    self.assertEqual(p['original_scaling']['E'],0.)
                    self.assertEqual(p['original_scaling']['G'],0.)
                    self.assertEqual(p['original_scaling']['b'],0.)

    def test_full_finite_original_dual_is_required(self):
        m,x,z=self.source('lower')
        for invalid in [np.array([1.]),np.full_like(z,np.nan),np.zeros_like(z)]:
            with self.subTest(size=len(invalid)),self.assertRaises(ValueError):
                transport.signed_transport(m,1.,x,invalid,'lower')

    def test_friction_and_fixed_loads_are_rejected(self):
        m,x,z=self.source('lower')
        m.materials[:,1:3]=1.
        with self.assertRaises(ValueError):transport.signed_transport(m,1.,x,z,'lower')
        m.materials[:,1:3]=0.;m.body[0,0,0]=1.
        with self.assertRaises(ValueError):transport.signed_transport(m,1.,x,z,'lower')

    def test_matrix_scaling_drift_is_rejected(self):
        m,x,z=self.source('lower');source=core.Prepared(m,1.,'lower',8)
        fresh=core.Prepared(m,2.,'lower',8);fresh.original[0].data[0]+=.01
        with self.assertRaisesRegex(ValueError,'homogeneity check failed'):
            transport.verify_scaling(source,fresh)


if __name__=='__main__':unittest.main()
