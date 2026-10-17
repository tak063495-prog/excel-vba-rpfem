"""Original physics and signed roots for the optional homogeneous c0 wedge."""
import copy
import math
import tempfile
import unittest
from argparse import Namespace
from pathlib import Path
import accuracy_driver as driver
import experiment_core as core
import c0_slope_wedge_driver as wedge
from report_accuracy import root_is_complete


class C0SlopeWedgeTests(unittest.TestCase):
    def setUp(self):
        self.parent, self.record = driver.load_model('D02A')

    def test_original_domain_material_body_and_exterior_are_preserved(self):
        mesh, angle, change = wedge.wedge_mesh(self.record, .01)
        self.assertEqual(mesh.physical_model_id, self.parent.physical_model_id)
        self.assertEqual(len(mesh.tri), 8)
        core.np.testing.assert_array_equal(mesh.xy[:6], self.record['source_case']['outline'])
        core.np.testing.assert_array_equal(mesh.materials, self.parent.materials)
        core.np.testing.assert_array_equal(mesh.body[:,:,0], 0.)
        core.np.testing.assert_array_equal(mesh.body[:,1,1], -20.)
        self.assertAlmostEqual(sum(core.topology(mesh)[0]), 1400.)
        self.assertAlmostEqual(math.tan(angle), 19.99/30)
        self.assertEqual(len(mesh.boundary), 6)
        self.assertTrue(change['same_physical_model'])

    def test_explicit_candidate_requires_the_original_physical_gate(self):
        mesh, angle, _ = wedge.wedge_mesh(self.record, .01)
        fs = 30/19.99+5e-5
        field, audit = wedge.explicit_witness(mesh, angle, fs)
        self.assertTrue(audit['physical_pass'])
        self.assertTrue(audit['witness'])
        self.assertFalse(audit['numerical_gate'])
        self.assertEqual(audit['q'], 8)
        field[0] += .1
        corrupted = core.evaluate(core.Prepared(mesh, fs, 'upper', 8, reduction='none'), field)
        self.assertFalse(corrupted['physical_pass'])
        self.assertFalse(corrupted['witness'])

    def test_unsupported_physics_and_invalid_geometry_are_rejected(self):
        for modify in [lambda r: r['materials'][0].__setitem__(1, .1),
                       lambda r: r['source_case'].__setitem__('pore', True),
                       lambda r: r['boundary_explicit'][0]['load0'].__setitem__(0, .1),
                       lambda r: r['settings'].__setitem__('SEISMIC_KH', .1),
                       lambda r: r['source_case']['outline'][3].__setitem__(1, 11.)]:
            record = copy.deepcopy(self.record);modify(record)
            with self.assertRaises(ValueError):wedge.wedge_mesh(record, .01)
        for depth in [0, -1, 20, float('nan'), float('inf')]:
            with self.assertRaises(ValueError):wedge.wedge_mesh(self.record, depth)
        _, mixed = driver.load_model('P06A')
        with self.assertRaises(ValueError):wedge.wedge_mesh(mixed, .01)

    def test_associated_and_davis_roots_close_on_their_own_original_meshes(self):
        with tempfile.TemporaryDirectory() as directory:
            for case in ['D02A', 'D03N']:
                report = wedge.run(Namespace(case=case, depth=.01, output=Path(directory),
                    tolerance=1e-4, solve_seconds=30.))
                metadata = report['upper'];_, record = driver.load_model(case)
                mesh = driver.load_saved_mesh(Path(directory)/case/'best_upper_mesh.npz', record)
                self.assertTrue(root_is_complete(metadata, mesh, metadata['fs'], 'upper', case))
                self.assertTrue(metadata['audit']['physical_pass'])
                self.assertTrue(metadata['audit']['witness'])
                self.assertFalse(report['mesh_change']['analytical_threshold_is_a_published_bound'])
                expected = 1.5 if case == 'D02A' else math.sqrt(1.25)
                self.assertGreater(metadata['fs'], expected)
                self.assertLess(metadata['fs']-expected, .001)


if __name__ == '__main__':unittest.main()
