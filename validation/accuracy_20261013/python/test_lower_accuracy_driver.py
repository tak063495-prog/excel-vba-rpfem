"""Own-mesh ancestry and positive-witness guards for the lower-only CLI."""
from argparse import Namespace
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
import accuracy_driver as driver
import experiment_core as core
from experiment_core import np
from lower_accuracy_driver import refine_lower, run
from mesh_accuracy import angular_fan_refine, lower_stress_score


class LowerDriverTests(unittest.TestCase):
    def test_two_level_parent_map_matches_sequential_polynomial_prolongation(self):
        mesh, record = core.load_case('P06A')
        field = np.r_[np.sin(np.arange(18*len(mesh.tri))/11), 1.000003]
        child, _, _ = refine_lower(mesh, record, field, .8, 'stress_fan', .6)
        first, _ = core.local_refine(mesh, record, lower_stress_score(mesh, field, .8), .6)
        second, _ = angular_fan_refine(first, record)
        self.assertEqual(child.mesh_id, second.mesh_id)
        direct = core.prolong(mesh, child, field, 'lower')
        sequential = core.prolong(first, second, core.prolong(mesh, first, field, 'lower'), 'lower')
        np.testing.assert_allclose(direct, sequential, rtol=1e-11, atol=1e-10)

    def test_a_zero_load_physical_field_cannot_be_an_imported_positive_bound(self):
        mesh, _ = driver.load_model('P06A')
        with TemporaryDirectory(prefix='rpfem_lower_cli_test_') as directory:
            folder = Path(directory)
            seed = folder/'seed'
            seed.mkdir()
            fs = .8
            driver.save_mesh(seed/'best_lower_mesh.npz', mesh)
            np.savez_compressed(seed/'best_lower_field.npz', x=np.zeros(18*len(mesh.tri)+1), fs=fs, mode='lower', q=8)
            core.write(seed/'import_audit.json', {'case':'P06A','certificates':{'lower':{
                'fs':fs,'mesh_id':mesh.mesh_id,'physical_model_id':mesh.physical_model_id}}})
            args = Namespace(case='P06A',seed=seed,method='uniform',output=folder/'output',rounds=0,
                             max_elements=24000,theta=.6,proximity_background=.1,tolerance=1e-4,solve_seconds=120,initial_span=.002)
            with self.assertRaisesRegex(ValueError, 'physical gate rejected'):
                run(args)
            self.assertFalse((folder/'output').exists())

    def test_initial_span_must_be_finite_and_positive_before_reading_a_seed(self):
        with TemporaryDirectory(prefix='rpfem_initial_span_test_') as directory:
            folder=Path(directory)
            for span in (0.,-1.,np.nan,np.inf):
                with self.subTest(span=span):
                    args=Namespace(case='P06A',seed=folder/'absent_seed',method='uniform',
                                   output=folder/'output',rounds=1,max_elements=24000,theta=.6,
                                   proximity_background=.1,tolerance=1e-4,solve_seconds=120,initial_span=span)
                    with self.assertRaisesRegex(ValueError,'initial span'):
                        run(args)
                    self.assertFalse((folder/'output').exists())


if __name__ == '__main__':
    unittest.main()
