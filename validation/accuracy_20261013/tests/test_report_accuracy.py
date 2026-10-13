"""Rollup control boundaries using synthetic metadata and temporary files.

The physical audit is isolated in control tests. The separate local P08
experiment checks the real uncondensed witness audit and copied evidence.
"""
from copy import deepcopy
import json
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import report_accuracy as report
from experiment_core import np,write


def mesh():
    return SimpleNamespace(xy=np.array([[0.,0.],[1.,0.],[0.,1.]]),
        tri=np.array([[0,1,2]],dtype=np.int32),material_id=np.array([0],dtype=np.int32),
        mesh_id='mesh-A',physical_model_id='model-A',
        provenance={'case':'P06A','psi_policy':'equal_reduced_phi'})


def metadata(mode='lower',fs=1.):
    m=mesh();bracket=[fs,fs+5e-5] if mode=='lower' else [fs-5e-5,fs]
    root={'mode':mode,'q':8,'case':'P06A','policy':m.provenance['psi_policy'],
        'mesh_id':m.mesh_id,'physical_model_id':m.physical_model_id,'root_complete':True,
        'endpoint':fs,'bracket':bracket,'bracket_signs':[1,-1]}
    return {'case':'P06A','mode':mode,'q':8,'fs':fs,'mesh_id':m.mesh_id,
        'physical_model_id':m.physical_model_id,'root_complete':True,
        'audit':{'discrete_root_search':root}}


class ReportBoundaryTests(unittest.TestCase):
    def test_strict_root_identity_and_bounds(self):
        for mode in ['lower','upper']:
            valid=metadata(mode);m=mesh()
            self.assertTrue(report.root_is_complete(valid,m,1.,mode,'P06A'))
            for key,value in [('case','P06N'),('mode','invalid'),('q',4),('mesh_id','other'),
                ('physical_model_id','other'),('fs',1.001),('root_complete',False)]:
                d=deepcopy(valid);d[key]=value
                with self.subTest(mode=mode,outer=key):
                    self.assertFalse(report.root_is_complete(d,m,1.,mode,'P06A'))
            for key,value in [('case','P06N'),('mode','invalid'),('q',4),('mesh_id','other'),
                ('physical_model_id','other'),('policy','other'),('endpoint',1.001),
                ('root_complete',False),('bracket_signs',[-1,1]),('bracket',[1.,1.01]),
                ('bracket',[1.01,1.]),('bracket',[0.,5e-5]),('bracket',[1.,np.nan]),
                ('bracket',[1.,np.inf]),('bracket',[])]:
                d=deepcopy(valid);d['audit']['discrete_root_search'][key]=value
                with self.subTest(mode=mode,root=key,value=value):
                    self.assertFalse(report.root_is_complete(d,m,1.,mode,'P06A'))

    def fixture(self,folder,document=None):
        catalog=folder/'catalog';catalog.mkdir();rec={'case':'P06A','source_case':{'pore':False}}
        write(catalog/'P06A.json',rec)
        case=folder/'P06A';case.mkdir();m=mesh()
        report.driver.save_mesh(case/'best_lower_mesh.npz',m)
        np.savez_compressed(case/'best_lower_field.npz',x=np.array([2.]),fs=1.,mode='lower',q=8)
        write(case/'import_audit.json',document or {'case':'P06A','certificates':{'lower':metadata()}})
        return catalog,rec,case/'import_audit.json',m

    def test_lower_only_import_audit_is_collected(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder=Path(tmp);catalog,rec,path,m=self.fixture(folder)
            with patch.object(report.driver,'CATALOG',catalog),\
                 patch.object(report.driver.core,'mesh_from_arrays',return_value=m),\
                 patch.object(report,'Prepared'),\
                 patch.object(report,'evaluate',return_value={'physical_pass':True,'witness':True}) as audit:
                lower=report.collect_candidate(path,'lower','P06A',rec,[folder],1e-4)
                upper=report.collect_candidate(path,'upper','P06A',rec,[folder],1e-4)
            self.assertTrue(lower['root_complete']);self.assertIsNone(upper)
            self.assertEqual(audit.call_count,1);self.assertEqual(report.case_of(path),'P06A')

    def test_metadata_mismatch_rejects_before_auditor(self):
        for key,value in [('fs',1.01),('mesh_id','other'),('physical_model_id','other')]:
            with self.subTest(key=key),tempfile.TemporaryDirectory() as tmp:
                folder=Path(tmp);d=metadata();d[key]=value
                catalog,rec,path,m=self.fixture(folder,{'certificates':{'lower':d}})
                with patch.object(report.driver,'CATALOG',catalog),\
                     patch.object(report.driver.core,'mesh_from_arrays',return_value=m),\
                     patch.object(report,'evaluate') as audit,self.assertRaisesRegex(ValueError,'inconsistent'):
                    report.collect_candidate(path,'lower','P06A',rec,[folder],1e-4)
                audit.assert_not_called()

    def test_mutation_during_audit_rejects_snapshot(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder=Path(tmp);catalog,rec,path,m=self.fixture(folder)
            def audit(*args):
                np.savez_compressed(path.parent/'best_lower_field.npz',x=np.array([3.]),fs=1.,mode='lower',q=8)
                return {'physical_pass':True,'witness':True}
            with patch.object(report.driver,'CATALOG',catalog),\
                 patch.object(report.driver.core,'mesh_from_arrays',return_value=m),\
                 patch.object(report,'Prepared'),patch.object(report,'evaluate',side_effect=audit),\
                 self.assertRaisesRegex(ValueError,'MUTABLE_SOURCE_CHANGED_DURING_REAUDIT'):
                report.collect_candidate(path,'lower','P06A',rec,[folder],1e-4)

    def test_own_copy_survives_source_change_and_stale_freeze_rejects(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder=Path(tmp);source=folder/'source.txt';source.write_bytes(b'original');tracked={}
            report.read_tracked(source,tracked)
            candidate={'mode':'lower','fs':1.,'mesh':mesh(),'field':np.array([2.]),
                'audit':{'physical_pass':True,'witness':True,'discrete_root_search':metadata()['audit']['discrete_root_search']},
                'source':str(source),'root_complete':True,'tracked':tracked,
                'source_hashes':{str(p):h for p,h in tracked.items()}}
            copied=folder/'snapshots/P06A';published=report.freeze_candidate(candidate,copied)
            ownfield=copied/published['field_path'];frozen_bytes=ownfield.read_bytes()
            source.write_bytes(b'changed');candidate['field'][0]=9.
            self.assertEqual(ownfield.read_bytes(),frozen_bytes)
            with np.load(ownfield,allow_pickle=False) as d:self.assertEqual(float(d['x'][0]),2.)
            with self.assertRaisesRegex(ValueError,'MUTABLE_SOURCE_CHANGED_BEFORE_FREEZE'):
                report.freeze_candidate(candidate,folder/'second')
            self.assertFalse((folder/'second').exists())

    def test_output_excluded_from_recursive_sources_and_path_escape_rejected(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder=Path(tmp);output=folder/'rollup'
            write(folder/'input/summary.json',{'case':'P06A'})
            write(output/'snapshots/P06A/summary.json',{'case':'P06A'})
            self.assertEqual(report.source_files([folder],output),[folder/'input/summary.json'])
            with self.assertRaises(ValueError):report.inside_file(folder/'input','../../outside.npz')


if __name__=='__main__':unittest.main()
