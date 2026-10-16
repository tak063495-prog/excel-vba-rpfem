"""Original physical equation multipliers without changing the six-output API."""
import unittest
import copy
import pickle
import accuracy_driver as driver
import experiment_core as core
from experiment_core import np
from model_support.lower_assembly import Program,assemble_lower


class LowerRowNormTests(unittest.TestCase):
    def test_real_free_and_fixed_load_programs_keep_tuple_contract(self):
        for case in ['D01A','P06A','P06N']:
            mesh,_=driver.load_model(case)
            for fixed_load in [None,1.000003]:
                with self.subTest(case=case,fixed_load=fixed_load):
                    p=assemble_lower(mesh,.9,fixed_load)
                    E,b,G,h,dims,zero=p
                    self.assertIsInstance(p,tuple)
                    self.assertEqual(len(p),6)
                    self.assertIs(p.eq,E)
                    self.assertIs(p.rhs,b)
                    self.assertIs(p.G,G)
                    self.assertIs(p.zero,zero)
                    self.assertEqual(p.row_norm.shape,(len(b),))
                    self.assertTrue(np.isfinite(p.row_norm).all())
                    self.assertTrue(np.all(p.row_norm>0))
                    self.assertFalse(p.row_norm.flags.writeable)
                    np.testing.assert_allclose(np.asarray(E.multiply(E).sum(axis=1)).ravel(),1.,rtol=5e-15,atol=5e-15)
                    if fixed_load is not None:
                        self.assertEqual(p.row_norm[-1],1.)
                        self.assertEqual(b[-1],fixed_load)
                        self.assertEqual(E[-1,-1],1.)

    def test_weights_restore_analytic_stress_equilibrium_in_every_material(self):
        # A linear stress polynomial has exact constant divergence. The
        # product row_norm*(E*x-b) must recover that physical equation,
        # including fixed and reference body loads at all Bernstein vertices.
        for case in ['D01A','P06A','P06N']:
            mesh,_=driver.load_model(case)
            p=assemble_lower(mesh,.9)
            points=mesh.xy[mesh.tri]
            points=np.concatenate([points,(points+np.roll(points,-1,axis=1))/2],axis=1)
            x,y=points[:,:,0],points[:,:,1]
            stress=np.stack([2*x+3*y+1,5*x-7*y-4,-.5*x+.7*y+.8],axis=-1)
            load=.73
            field=np.r_[stress.ravel()/mesh.stress_scale,load]
            residual=p.row_norm*(p.eq@field-p.rhs)*mesh.stress_scale
            cursor=0
            for e in range(len(mesh.tri)):
                expected=np.array([2.7,-7.5])+mesh.body[e,:,0]+load*mesh.body[e,:,1]
                np.testing.assert_allclose(residual[cursor:cursor+6],np.tile(expected,3),rtol=1e-10,atol=1e-9)
                cursor+=6
                count=3*int(np.sum(p.zero[e]))
                if count:
                    np.testing.assert_allclose(residual[cursor:cursor+count],stress[e,p.zero[e]].ravel(),rtol=5e-15,atol=1e-12)
                cursor+=count

    def test_invalid_metadata_cannot_supply_a_correction_preconditioner(self):
        mesh,_=driver.load_model('P06A')
        p=assemble_lower(mesh,.9)
        for bad in [p.row_norm[:-1],p.row_norm*np.nan,np.zeros(len(p.row_norm))]:
            with self.assertRaises(ValueError):Program(*p,bad)

    def test_tuple_serialization_preserves_readonly_metadata(self):
        mesh,_=driver.load_model('P06A')
        source=assemble_lower(mesh,.9)
        for candidate in [copy.deepcopy(source),pickle.loads(pickle.dumps(source))]:
            self.assertIs(candidate.eq,candidate[0])
            self.assertIs(candidate.zero,candidate[5])
            self.assertFalse(candidate.row_norm.flags.writeable)
            np.testing.assert_array_equal(candidate.eq.data,source.eq.data)
            np.testing.assert_array_equal(candidate.row_norm,source.row_norm)


if __name__=='__main__':unittest.main()
