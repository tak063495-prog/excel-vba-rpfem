"""Light control-flow regressions, not numerical certificates or speed claims."""
from pathlib import Path
from types import SimpleNamespace
import argparse
import json
import sys
import tempfile
import unittest
from unittest.mock import patch

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import accuracy_driver as driver
import numpy as np


class Clock:
    def __init__(self):self.t=0.
    def __call__(self):return self.t
    def advance(self,seconds):self.t+=seconds


def mesh():
    return SimpleNamespace(xy=np.array([[0.,0.],[1.,0.],[0.,1.]]),
        tri=np.array([[0,1,2]],dtype=np.int32),material_id=np.array([0],dtype=np.int32),
        materials=np.array([[1.,.1,.1]]),body=np.zeros((1,2,2)),
        provenance={'case':'TEST','psi_policy':'equal_reduced_phi'},
        mesh_id='test-mesh',physical_model_id='test-physical-model')


def result(witness=False,value=0.,numerical=False):
    return dict(status='synthetic control fixture',physical_pass=witness,witness=witness,
                numerical_gate=numerical,value=value,end_to_end_seconds=0.)


def args(output,**overrides):
    values=dict(case='TEST',output=Path(output),seed=None,resume=False,recheck=False,
        fan_lower=False,upper_only=False,cycles=0,max_elements=100,theta=.6,
        tolerance=1e-4,solve_seconds=30.,case_seconds=900.,target_gap=.02)
    values.update(overrides)
    return argparse.Namespace(**values)


class BudgetTests(unittest.TestCase):
    def test_cap_and_shared_remaining(self):
        clock=Clock();budget=driver.CaseBudget(35,30,clock)
        self.assertEqual(budget.allowance('LP'),30)
        clock.advance(20)
        self.assertEqual(budget.allowance('dual',30),15)
        self.assertEqual(budget.allowance('helper',4),4)
        clock.advance(16)
        with self.assertRaises(driver.BudgetExceeded):budget.allowance('next')
        self.assertTrue(budget.record()['exhausted'])
        self.assertEqual(len(budget.record()['solver_time_allocations']),3)

    def test_invalid_limits(self):
        for seconds in [0.,-1.,float('nan'),float('inf')]:
            with self.assertRaises(ValueError):driver.CaseBudget(seconds,30)
            with self.assertRaises(ValueError):driver.CaseBudget(10,seconds)

    def test_expiry_preserves_new_physical_endpoint_without_root_promotion(self):
        clock=Clock();budget=driver.CaseBudget(1,30,clock);limits=[]
        def solve(prepared,options,**kwargs):
            limits.append(options['time_limit']);clock.advance(1.1)
            return result(True,1.01),np.ones(18),None
        refresh=SimpleNamespace(at=lambda fs:SimpleNamespace(original=None))
        with tempfile.TemporaryDirectory() as output,patch.object(driver,'LowerRefresh',return_value=refresh),patch.object(driver,'solve',side_effect=solve):
            certificate,summary=driver.search(mesh(),'lower',(.4,1.3),Path(output),'timed',budget=budget)
            self.assertEqual(limits,[1.])
            self.assertEqual(certificate.fs,.4)
            self.assertFalse(certificate.root_complete)
            self.assertTrue(summary['budget_exhausted'])
            self.assertEqual(summary['bracket_signs'],[1,None])
            with np.load(Path(output)/'best_lower_field.npz') as field:
                self.assertEqual(float(field['fs']),.4)
                self.assertTrue(np.all(field['x']==1))
            saved=json.loads((Path(output)/'best_lower_certificate.json').read_text())
            self.assertFalse(saved['root_complete'])

    def test_lp_retries_share_one_budget(self):
        clock=Clock();budget=driver.CaseBudget(1,30,clock);limits=[]
        def solve(prepared,options,**kwargs):
            limits.append(options['time_limit']);clock.advance(.6)
            return result(),np.zeros(18),None
        refresh=SimpleNamespace(at=lambda fs:SimpleNamespace(original=None))
        with tempfile.TemporaryDirectory() as output,patch.object(driver,'LowerRefresh',return_value=refresh),patch.object(driver,'solve',side_effect=solve):
            certificate,summary=driver.search(mesh(),'lower',(.4,1.3),Path(output),'retry',budget=budget)
        self.assertIsNone(certificate)
        self.assertEqual(len(limits),2)
        self.assertAlmostEqual(limits[1],.4)
        self.assertTrue(summary['budget_exhausted'])

    def test_original_helper_uses_remaining_case_time(self):
        import original_dual
        clock=Clock();budget=driver.CaseBudget(.5,30,clock);limits=[]
        def solve(prepared,options,**kwargs):
            limits.append(options['time_limit']);clock.advance(.1)
            return result(),np.zeros(12),None
        helper=[]
        def dual(prepared,seconds):
            helper.append(seconds);clock.advance(.2)
            return {'accepted':False}
        prepared=SimpleNamespace(original=(np.ones(1),))
        with tempfile.TemporaryDirectory() as output,patch.object(driver,'Prepared',return_value=prepared),patch.object(driver,'solve',side_effect=solve),patch.object(original_dual,'solve_original_dual',side_effect=dual):
            certificate,summary=driver.search(mesh(),'upper',(.4,1.3),Path(output),'helper',budget=budget)
        self.assertIsNone(certificate)
        self.assertEqual(len(limits),4)
        self.assertEqual(len(helper),1)
        self.assertAlmostEqual(helper[0],.1)
        self.assertTrue(summary['budget_exhausted'])

    def test_fixed_load_recovery_uses_remaining_case_time(self):
        clock=Clock();budget=driver.CaseBudget(.8,30,clock);limits=[]
        def solve(prepared,options,**kwargs):
            limits.append(options['time_limit']);clock.advance(.5)
            return result(False,1.00001,True),np.zeros(18),None
        refresh=SimpleNamespace(at=lambda fs:SimpleNamespace(original=None))
        with tempfile.TemporaryDirectory() as output,patch.object(driver,'LowerRefresh',return_value=refresh),patch.object(driver,'Prepared',return_value=SimpleNamespace()),patch.object(driver,'solve',side_effect=solve):
            certificate,summary=driver.search(mesh(),'lower',(.4,1.3),Path(output),'fixed',budget=budget)
        self.assertIsNone(certificate)
        self.assertEqual(len(limits),2)
        self.assertAlmostEqual(limits[1],.3)
        self.assertTrue(summary['budget_exhausted'])

    def run_synthetic_pair(self,gap,root=True,expire=False):
        clock=Clock();real_budget=driver.CaseBudget
        m=mesh()
        def search(m,mode,initial,folder,label,*pos,**kwargs):
            fs=1. if mode=='lower' else (2+gap)/(2-gap)
            cert=driver.Certificate(mode,fs,m,np.ones(18 if mode=='lower' else 12),
                result(True,1.01 if mode=='lower' else .99),'control fixture',root)
            if expire and mode=='upper':clock.advance(901)
            return cert,{'root_complete':root}
        with tempfile.TemporaryDirectory() as output,patch.object(driver,'load_model',return_value=(m,{})),patch.object(driver,'search',side_effect=search),patch.object(driver,'CaseBudget',side_effect=lambda total,per:real_budget(total,per,clock)):
            report=driver.run(args(output))
            self.assertTrue((Path(output)/'TEST'/'summary.json').is_file())
            self.assertTrue((Path(output)/'TEST'/'best_lower_field.npz').is_file())
            self.assertTrue((Path(output)/'TEST'/'best_upper_field.npz').is_file())
            return report

    def test_two_percent_accepts_1_99_percent(self):
        report=self.run_synthetic_pair(.0199)
        self.assertEqual(report['status'],'PASS')
        self.assertTrue(report['target_met'])
        self.assertFalse(report['one_percent_width_pass'])
        self.assertEqual(report['execution_scope']['kind'],'cold_start_new_python_policy')

    def test_two_percent_rejects_2_01_percent(self):
        report=self.run_synthetic_pair(.0201)
        self.assertEqual(report['status'],'PARTIAL')
        self.assertFalse(report['target_met'])

    def test_narrow_pair_requires_both_roots(self):
        report=self.run_synthetic_pair(.005,root=False)
        self.assertEqual(report['status'],'PARTIAL')
        self.assertFalse(report['root_complete'])
        self.assertFalse(report['target_met'])

    def test_expired_narrow_closed_pair_is_never_pass(self):
        report=self.run_synthetic_pair(.005,expire=True)
        self.assertEqual(report['status'],'BUDGET_EXHAUSTED')
        self.assertFalse(report['target_met'])
        self.assertIn('lower',report);self.assertIn('upper',report)

    def test_original_gates_and_source_binding_unchanged(self):
        self.assertEqual(driver.core.PHYSICAL_TOL,1e-7)
        self.assertEqual(driver.core.RAW_STRESS_TOL,1e-7)
        self.assertEqual(driver.core.LOWER_TARGET,1.000001)
        self.assertEqual(driver.core.UPPER_TARGET,.999999)
        sources=driver.source_hashes()
        self.assertTrue(any(key.endswith('/benchmark_models.py') for key in sources))
        self.assertTrue(all('/G1R19_python26_v5_20261016/python/' in key for key in sources))


if __name__=='__main__':unittest.main()
