"""Geometry and polynomial-witness checks for optional angular fan meshes."""
import unittest
from collections import Counter
import experiment_core as core
from experiment_core import np, topology
from mesh_accuracy import angular_fan_refine


def interface_edges(mesh):
    _, _, edges = topology(mesh)
    return {(e.a, e.b) for e in edges if e.elements[1] >= 0
            and mesh.material_id[e.elements[0]] != mesh.material_id[e.elements[1]]}


class AngularFanTests(unittest.TestCase):
    def test_boundary_interfaces_areas_and_parent_maps(self):
        for case in ('P06A', 'P06N'):
            with self.subTest(case=case):
                mesh, rec = core.load_case(case)
                # An existing parent mapping must not replace direct child IDs.
                mesh.provenance['parent_element_ids'] = [999] * len(mesh.tri)
                child, info = angular_fan_refine(mesh, rec)
                self.assertEqual(child.physical_model_id, mesh.physical_model_id)
                np.testing.assert_array_equal(child.xy[:len(mesh.xy)], mesh.xy)
                self.assertEqual(set(child.boundary), set(mesh.boundary))
                for key in mesh.boundary:
                    self.assertEqual(child.boundary[key]['kind'], mesh.boundary[key]['kind'])
                    np.testing.assert_array_equal(child.boundary[key]['loads'], mesh.boundary[key]['loads'])
                self.assertEqual(interface_edges(child), interface_edges(mesh))
                area, _, _ = topology(mesh)
                ca, _, _ = topology(child)
                parents = np.asarray(child.provenance['parent_element_ids'])
                self.assertEqual(len(parents), len(child.tri))
                self.assertTrue(np.all((0 <= parents) & (parents < len(mesh.tri))))
                np.testing.assert_allclose(np.bincount(parents, weights=ca, minlength=len(area)), area, rtol=1e-12, atol=1e-12)
                np.testing.assert_array_equal(child.material_id, mesh.material_id[parents])
                counts = Counter(parents)
                self.assertTrue(set(counts.values()) <= {1, 3})
                self.assertEqual(sum(v == 3 for v in counts.values()), info['split_elements'])
                self.assertEqual(len(child.tri), len(mesh.tri) + 2 * info['split_elements'])

    def test_p2_prolongation_recovers_same_physical_polynomial(self):
        mesh, rec = core.load_case('P06A')
        child, _ = angular_fan_refine(mesh, rec)
        # Quadratic displacement fields and quadratic Bernstein stress fields.
        velocity = np.arange(12 * len(mesh.tri), dtype=float) / 100
        stress = np.r_[np.arange(18 * len(mesh.tri), dtype=float) / 100, 1.000003]
        pv = core.prolong(mesh, child, velocity, 'upper').reshape(-1, 6, 2)
        ps = core.prolong(mesh, child, stress, 'lower').reshape(-1)[:-1].reshape(-1, 6, 3)
        np.testing.assert_allclose(core.prolong(mesh, child, stress, 'lower')[-1], stress[-1])
        ov = velocity.reshape(-1, 6, 2)
        os = stress[:-1].reshape(-1, 6, 3)
        for k, e in enumerate(child.provenance['parent_element_ids']):
            # Evaluate at a point away from interpolation nodes.
            cb = np.asarray([.2, .3, .5])
            point = cb @ child.xy[child.tri[k]]
            pb = np.linalg.solve(np.vstack([mesh.xy[mesh.tri[e]].T, np.ones(3)]), np.r_[point, 1.])
            def bernstein(b):
                a, c, d = b
                return np.array([a*a, c*c, d*d, 2*a*c, 2*c*d, 2*d*a])
            def nodal(b):
                a, c, d = b
                return np.array([a*(2*a-1), c*(2*c-1), d*(2*d-1), 4*a*c, 4*c*d, 4*d*a])
            np.testing.assert_allclose(nodal(cb) @ pv[k], nodal(pb) @ ov[e], rtol=1e-11, atol=1e-10)
            np.testing.assert_allclose(bernstein(cb) @ ps[k], bernstein(pb) @ os[e], rtol=1e-11, atol=1e-10)

    def test_scope_validation(self):
        mesh, rec = core.load_case('P06A')
        with self.assertRaises(ValueError):
            angular_fan_refine(mesh, rec, 'all')


if __name__ == '__main__':
    unittest.main()
