"""Protected geometry and polynomial transfer checks for non-nested flips."""
import unittest
from argparse import Namespace
from pathlib import Path
from tempfile import TemporaryDirectory
import experiment_core as core
from experiment_core import np
from mesh_geometry import flow_metric_flip, interpolate_velocity, audit_geometry, protected_edges
import accuracy_driver as driver
from upper_geometry_driver import run


def nodal_points(mesh):
    vertices=mesh.xy[mesh.tri]
    return np.concatenate([vertices,(vertices+np.roll(vertices,-1,axis=1))/2],axis=1)


def polynomial(points):
    x,y=points[...,0],points[...,1]
    return np.stack([x*x+.3*x*y+2*y,-x*y+.2*y*y+3*x],axis=-1)


class FlowGeometryTests(unittest.TestCase):
    def test_flips_preserve_all_exterior_material_edges_and_areas(self):
        for case in ('P06A','P06N'):
            with self.subTest(case=case):
                mesh,record=core.load_case(case)
                velocity=polynomial(nodal_points(mesh)).ravel()
                child,info=flow_metric_flip(mesh,record,velocity,passes=2)
                self.assertGreater(info['flips'],0)
                self.assertEqual(len(mesh.tri),len(child.tri))
                self.assertEqual(mesh.physical_model_id,child.physical_model_id)
                self.assertEqual(protected_edges(mesh),protected_edges(child))
                np.testing.assert_array_equal(mesh.xy,child.xy)
                self.assertFalse(info['certificate'])
                self.assertFalse(info['nested'])
                self.assertLess(info['geometry_audit']['maximum_material_area_difference'],1e-10)
                self.assertNotIn('parent_element_ids',child.provenance)

    def test_nonnested_transfer_recovers_quadratic_polynomials_without_crossing_material(self):
        mesh,record=core.load_case('P06A')
        velocity=polynomial(nodal_points(mesh)) + mesh.material_id[:,None,None]*10
        child,_=flow_metric_flip(mesh,record,velocity.ravel(),passes=2)
        result=interpolate_velocity(mesh,child,velocity.ravel())[:12*len(child.tri)].reshape(-1,6,2)
        expected=polynomial(nodal_points(child)) + child.material_id[:,None,None]*10
        np.testing.assert_allclose(result,expected,rtol=1e-11,atol=1e-10)

    def test_geometry_guard_rejects_moved_nodes_and_flow_input_is_finite(self):
        mesh,record=core.load_case('P06A')
        child=core.mesh_from_arrays(mesh.xy.copy(),mesh.tri.copy(),mesh.material_id.copy(),record)
        child.xy[0,0]+=.001
        with self.assertRaisesRegex(ValueError,'coordinates changed'):
            audit_geometry(mesh,child)
        velocity=np.zeros(12*len(mesh.tri))
        for invalid in [velocity[:-1],np.r_[np.nan,velocity[1:]]]:
            with self.assertRaises(ValueError):flow_metric_flip(mesh,record,invalid)
            with self.assertRaises(ValueError):interpolate_velocity(mesh,mesh,invalid)
        with self.assertRaises(ValueError):flow_metric_flip(mesh,record,velocity,passes=-1)
        with self.assertRaises(ValueError):flow_metric_flip(mesh,record,velocity,passes=1.5)

    def test_geometry_cli_rejects_unphysical_source_before_export(self):
        mesh,_=driver.load_model('P06A')
        with TemporaryDirectory(prefix='rpfem_upper_geometry_test_') as directory:
            folder=Path(directory);seed=folder/'seed';seed.mkdir()
            driver.save_mesh(seed/'best_upper_mesh.npz',mesh)
            field=core.complete_upper_velocity(mesh,np.zeros(12*len(mesh.tri)),8)
            np.savez_compressed(seed/'best_upper_field.npz',x=field,fs=.85,mode='upper',q=8)
            args=Namespace(case='P06A',seed=seed,output=folder/'output',passes=1,
                           bracket=None,tolerance=1e-4,solve_seconds=120)
            with self.assertRaisesRegex(ValueError,'full original physical gate'):
                run(args)
            self.assertFalse((folder/'output').exists())


if __name__=='__main__':unittest.main()
