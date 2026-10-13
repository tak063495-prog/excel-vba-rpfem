"""Cached original dual vectors guide a root without creating a primal bound."""
from pathlib import Path
import copy
import hashlib
import json
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import accuracy_driver as driver
import c0_slope_wedge_driver as wedge
import experiment_core as core
import original_dual
from test_original_dual import lower


def guide_for(mesh, mode, fs, dual):
    return {'case':mesh.provenance['case'], 'mode':mode, 'q':8, 'fs':fs,
        'mesh_id':mesh.mesh_id, 'physical_model_id':mesh.physical_model_id,
        'policy':mesh.provenance['psi_policy'], 'dual':dual}


class InitialDualGuideTests(unittest.TestCase):
    def specimen(self):
        prepared=lower(.5);prepared.m.provenance['psi_policy']='equal_reduced_phi'
        return prepared, guide_for(prepared.m, 'lower', 1., core.np.array([1.]))

    def test_raw_reaudit_creates_only_an_inverse_sign_and_keeps_primal_unchanged(self):
        prepared,guide=self.specimen();field=prepared.physical_x.copy()
        with patch.object(driver,'Prepared',return_value=prepared):
            fs,result,dual=driver.audit_initial_dual_guide(prepared.m,'lower',guide)
        self.assertEqual(fs,1.);self.assertEqual(driver.sign(result,'lower'),-1)
        self.assertFalse(result['physical_pass']);self.assertFalse(result['witness'])
        self.assertFalse(result['publishes_primal_bound']);self.assertTrue(result['dual_guide_only'])
        core.np.testing.assert_array_equal(prepared.physical_x,field)
        core.np.testing.assert_array_equal(dual,guide['dual'])

    def test_metadata_mismatch_is_rejected_before_reassembly(self):
        prepared,guide=self.specimen()
        for key,value in [('case','other'),('mode','upper'),('q',4),('mesh_id','other'),
                          ('physical_model_id','other'),('policy','other')]:
            modified=dict(guide,**{key:value})
            with patch.object(driver,'Prepared') as assembler:
                with self.assertRaises(ValueError):driver.audit_initial_dual_guide(prepared.m,'lower',modified)
                assembler.assert_not_called()
        for fs in [0.,-1.,float('nan'),float('inf'),'1.']:
            with patch.object(driver,'Prepared') as assembler:
                with self.assertRaises(ValueError):driver.audit_initial_dual_guide(prepared.m,'lower',dict(guide,fs=fs))
                assembler.assert_not_called()

    def test_bad_original_vectors_and_fixed_load_cost_are_rejected(self):
        prepared,guide=self.specimen()
        for dual in [core.np.array([core.np.nan]),core.np.array([core.np.inf]),
                     core.np.array([1.01]),core.np.array([-1.]),core.np.array([])]:
            with patch.object(driver,'Prepared',return_value=prepared):
                with self.assertRaises(ValueError):driver.audit_initial_dual_guide(prepared.m,'lower',dict(guide,dual=dual))
        prepared.original[-1][-1]=0.
        with patch.object(driver,'Prepared',return_value=prepared):
            with self.assertRaises(ValueError):driver.audit_initial_dual_guide(prepared.m,'lower',guide)

    def test_original_guide_is_persisted_and_only_a_physical_witness_closes_a_real_root(self):
        _,record=driver.load_model('D02A');mesh,angle,_=wedge.wedge_mesh(record,.01)
        fs=1.45;prepared=core.Prepared(mesh,fs,'upper',8,reduction='none')
        evidence=original_dual.solve_original_dual(prepared,seconds=10.,return_dual=True)
        self.assertTrue(evidence['accepted'])
        guide=guide_for(mesh,'upper',fs,evidence['dual'])
        warm_fs=30/19.99+5e-5;x,audit=wedge.explicit_witness(mesh,angle,warm_fs)
        warm=driver.Certificate('upper',warm_fs,mesh,x,audit,'original physical wedge',False)
        with tempfile.TemporaryDirectory() as directory:
            folder=Path(directory)
            best,root=driver.search(mesh,'upper',(fs,warm_fs),folder,'guide_root',
                1e-4,30,warm,initial_dual_guides=[guide])
            self.assertTrue(root['root_complete']);self.assertTrue(best.result['witness'])
            self.assertFalse(best.result.get('dual_guide_only',False))
            rows=json.loads((folder/'guide_root_trials.json').read_text(encoding='utf8'))
            cached=next(row for row in rows if row.get('dual_guide_only'))
            self.assertFalse(cached['physical_pass']);self.assertFalse(cached['witness'])
            proof=folder/cached['dual_proof_path']
            self.assertEqual(hashlib.sha256(proof.read_bytes()).hexdigest(),cached['dual_proof_sha256'])
            with core.np.load(proof,allow_pickle=False) as saved:
                core.np.testing.assert_array_equal(saved['dual'],guide['dual'])
                self.assertNotIn('x',saved.files)

    def test_conflicting_cached_primal_and_dual_are_rejected(self):
        prepared,guide=self.specimen()
        certificate=driver.Certificate('lower',1.,prepared.m,prepared.physical_x.copy(),
            {'witness':True,'physical_pass':True,'value':1.000001},'unchanged witness',False)
        with tempfile.TemporaryDirectory() as directory:
            with patch.object(driver,'LowerRefresh'),patch.object(driver,'Prepared',return_value=prepared):
                with self.assertRaises(ValueError):
                    driver.search(prepared.m,'lower',(.9,1.1),Path(directory),'conflict',
                        initial_witness=certificate,initial_dual_guides=[guide])


if __name__=='__main__':unittest.main()
