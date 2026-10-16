"""Control regressions for explicit incomplete pilots and strict terminal work."""
from pathlib import Path
from types import SimpleNamespace
import importlib.util
import json
import sys
import tempfile
import unittest
from unittest.mock import patch
import numpy as np

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'python'))
import accuracy_driver as driver
from test_accuracy_budget import args,mesh,result


class PilotTests(unittest.TestCase):
    def search(self,limit,bracket=(.4,1.3)):
        def solve(prepared,options,**kwargs):
            positive=prepared.fs<.8
            r=result(positive,1.01 if positive else .99,True)
            r['physical_pass']=True
            return r,np.ones(19),None
        refresh=SimpleNamespace(at=lambda fs:SimpleNamespace(original=None,fs=fs))
        with tempfile.TemporaryDirectory() as out,patch.object(driver,'LowerRefresh',return_value=refresh),\
                patch.object(driver,'solve',side_effect=solve):
            cert,summary=driver.search(mesh(),'lower',bracket,Path(out),'pilot',max_trials=limit,bisection=True)
            saved=json.loads((Path(out)/'best_lower_certificate.json').read_text())
            return cert,summary,saved

    def test_one_trial_preserves_field_but_never_claims_root(self):
        cert,summary,saved=self.search(1)
        self.assertEqual(summary['new_points'],1)
        self.assertTrue(summary['pilot_trial_limit_reached'])
        self.assertFalse(summary['budget_exhausted'])
        self.assertFalse(summary['root_complete']);self.assertFalse(saved['root_complete'])
        self.assertEqual(cert.fs,.4)

    def test_cached_initial_witness_does_not_spend_point_budget(self):
        m=mesh();warm=driver.Certificate('lower',.5,m,np.ones(19),result(True,1.01,True),'control')
        refresh=SimpleNamespace(at=lambda fs:SimpleNamespace(original=None,fs=fs))
        with tempfile.TemporaryDirectory() as out,patch.object(driver,'LowerRefresh',return_value=refresh),\
                patch.object(driver,'solve',return_value=(result(True,1.01,True),np.ones(19),None)) as solve:
            cert,summary=driver.search(m,'lower',(.4,1.3),Path(out),'cached',initial_witness=warm,max_trials=1)
        self.assertEqual(solve.call_count,1)
        self.assertEqual(summary['new_points'],1)
        self.assertFalse(cert.root_complete)

    def test_strict_width_can_complete_naturally_under_point_cap(self):
        cert,summary,_=self.search(2,(.79999,.80001))
        self.assertTrue(cert.root_complete)
        self.assertEqual(summary['bracket_signs'],[1,-1])
        self.assertLessEqual(summary['bracket_width'],1e-4)

    def test_unlimited_default_retains_original_root_contract(self):
        cert,summary,_=self.search(None)
        self.assertTrue(cert.root_complete)
        self.assertIsNone(summary['pilot_trial_limit'])
        self.assertGreater(summary['new_points'],4)

    def test_invalid_caps(self):
        for value in (0,-1,1.5,True,129):
            with tempfile.TemporaryDirectory() as out,self.assertRaises(ValueError):
                driver.search(mesh(),'lower',(.4,1.3),Path(out),'invalid',max_trials=value)

    def terminal_pair(self,complete_terminal):
        calls=[];m=mesh()
        def search(mesh,mode,initial,folder,label,*pos,**kwargs):
            strict=kwargs.get('max_trials') is None
            calls.append((mode,strict))
            root=bool(strict and complete_terminal)
            fs=1. if mode=='lower' else 1.01
            cert=driver.Certificate(mode,fs,m,np.ones(19 if mode=='lower' else 12),
                result(True,1.01 if mode=='lower' else .99),'control only',root)
            return cert,dict(root_complete=root,max_trials=kwargs.get('max_trials'))
        with tempfile.TemporaryDirectory() as out,patch.object(driver,'load_model',return_value=(m,{})),\
                patch.object(driver,'search',side_effect=search):
            report=driver.run(args(out,pilot_trials=4))
        return calls,report

    def test_physical_two_percent_pair_requires_both_unlimited_terminal_roots(self):
        calls,report=self.terminal_pair(True)
        self.assertEqual(calls,[('lower',False),('upper',False),('lower',True),('upper',True)])
        self.assertTrue(report['target_met']);self.assertEqual(report['status'],'PASS')
        self.assertEqual(len(report['terminal_root_closures']),2)

    def test_failed_terminal_roots_never_promote_narrow_physical_pair(self):
        calls,report=self.terminal_pair(False)
        self.assertFalse(report['target_met']);self.assertFalse(report['root_complete'])
        self.assertEqual(report['status'],'PARTIAL')

    def test_owned_stop_is_child_first_and_stale_identity_not_killed(self):
        spec=importlib.util.spec_from_file_location('v5_supervision_under_test',ROOT/'supervise_p06.py')
        support=importlib.util.module_from_spec(spec);spec.loader.exec_module(support)
        order=[]
        class FakeProcess:
            def __init__(self,pid):self.pid=pid
            def kill(self):order.append(self.pid)
        with tempfile.TemporaryDirectory() as out:
            root=Path(out);folder=root/'fields';folder.mkdir();(folder/'proof.json').write_text('{"witness":true}')
            job=dict(case='TEST',parent=dict(pid=10),descendants={'11':dict(pid=11),'12':dict(pid=12)},
                folder=str(folder),ownership_path=str(root/'TEST_OWNERSHIP.json'))
            with patch.object(support,'discover'),patch.object(support,'same_process',side_effect=lambda r:r['pid']!=12),\
                    patch.object(support.psutil,'Process',FakeProcess):
                support.terminate_owned(job,'control test')
            self.assertEqual(order,[11,10])
            self.assertFalse(job['interruption']['new_bound_adopted'])
            self.assertEqual(job['interruption']['checkpoint_sha256_before'],job['interruption']['checkpoint_sha256_after'])


if __name__=='__main__':unittest.main()
