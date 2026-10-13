"""Regression of certificate boundaries, using temporary synthetic evidence only.

No production meshes, fields, result files or solver runs are changed. Physical
admissibility is controlled at the audit boundary in metadata/control tests;
the numerical experiments perform the independent real-model physical audits.
"""
from __future__ import annotations

from contextlib import redirect_stdout
from copy import deepcopy
import io
import json
from pathlib import Path
import sys
import tempfile
from types import SimpleNamespace
import unittest
from unittest.mock import patch

PYTHON = Path(__file__).resolve().parents[1] / "python"
sys.path.insert(0, str(PYTHON))
import accuracy_driver as driver
import experiment_core as core
from reduction import cone_error
from experiment_core import np, sparse


def mesh(mode="upper", physical="physical-A"):
    return SimpleNamespace(
        xy=np.array([[0., 0.], [1., 0.], [0., 1.]]),
        tri=np.array([[0, 1, 2]], dtype=np.int32),
        material_id=np.array([0], dtype=np.int32),
        materials=np.array([[1., 30., 30., 18.]]),
        body=np.zeros((1, 2, 2)),
        mesh_id="mesh-" + mode,
        physical_model_id=physical,
        provenance={"case": "P06A", "psi_policy": "ASSOCIATED"},
        stress_scale=1.,
    )


def result(value=1.000003, witness=True, physical=True, numerical=False, **extra):
    data = {"value": value, "physical_pass": physical, "witness": witness,
            "numerical_gate": numerical, "status": "Synthetic", "end_to_end_seconds": 0.}
    data.update(extra)
    return data


def root_metadata(mode, fs, m):
    bracket = [fs, fs + 5e-5] if mode == "lower" else [fs - 5e-5, fs]
    root = {"mode": mode, "q": 8, "case": "P06A", "policy": "ASSOCIATED",
            "mesh_id": m.mesh_id, "physical_model_id": m.physical_model_id,
            "root_complete": True, "endpoint": fs, "bracket": bracket,
            "bracket_width": bracket[1] - bracket[0], "bracket_signs": [1, -1],
            "fs_tolerance": 1e-4}
    return {"mode": mode, "q": 8, "case": "P06A", "fs": fs,
            "mesh_id": m.mesh_id, "physical_model_id": m.physical_model_id,
            "root_complete": True, "audit": {"discrete_root_search": root}}


class FiniteBoundaryTests(unittest.TestCase):
    def test_cone_error_rejects_every_nonfinite_component(self):
        for bad in (np.nan, np.inf, -np.inf):
            for dimension in (1, 3):
                for at in range(dimension):
                    values = np.zeros(dimension)
                    values[at] = bad
                    with self.subTest(bad=bad, dimension=dimension, at=at):
                        self.assertEqual(cone_error(values, [dimension]), np.inf)

    def test_nonfinite_raw_dual_never_supplies_search_sign(self):
        for mode in ("lower", "upper"):
            dual = -.99 if mode == "lower" else 1.01
            for key in ("dual_stationarity", "dual_cone_error", "dual_objective"):
                for bad in (np.nan, np.inf, -np.inf, None):
                    data = result(witness=False, physical=False,
                                  dual_stationarity=0., dual_cone_error=0., dual_objective=dual)
                    data[key] = bad
                    with self.subTest(mode=mode, key=key, bad=bad):
                        self.assertIsNone(driver.sign(data, mode))

    def test_reduced_dual_is_diagnostic_only(self):
        for mode in ("lower", "upper"):
            data = result(witness=False, physical=False, reduced_numerical_gate=True,
                          reduced_dual_stationarity=0., reduced_dual_cone_error=0.,
                          reduced_dual_objective=-.99 if mode == "lower" else 1.01)
            self.assertIsNone(driver.sign(data, mode))

    def test_admissible_primal_and_original_dual_have_separate_roles(self):
        self.assertEqual(driver.sign(result(), "lower"), 1)
        self.assertEqual(driver.sign(result(value=.999998), "upper"), -1)
        self.assertIsNone(driver.sign(result(value=np.nan), "lower"))
        self.assertIsNone(driver.sign(result(physical=False), "lower"))
        self.assertEqual(driver.sign(result(witness=False, physical=False,
                        dual_stationarity=0., dual_cone_error=0., dual_objective=-.99), "lower"), -1)
        self.assertEqual(driver.sign(result(witness=False, physical=False,
                        dual_stationarity=0., dual_cone_error=0., dual_objective=1.01), "upper"), 1)

    def prepared(self):
        m = mesh("lower")
        E = sparse.eye(1, format="csc")
        b = np.array([1.000003])
        G = sparse.csc_matrix((1, 1))
        return SimpleNamespace(original=(E, b, G, np.ones(1), [1], np.array([-1.])),
                               m=m, mode="lower", fs=1., q=8)

    def test_nonfinite_primal_returns_a_complete_unknown_schema(self):
        with patch.object(core, "audit") as audit:
            data = core.evaluate(self.prepared(), np.array([np.nan]))
        audit.assert_not_called()
        self.assertFalse(data["physical_pass"])
        self.assertFalse(data["witness"])
        self.assertFalse(data["numerical_gate"])
        self.assertIn("value", data)
        self.assertIsNone(driver.sign(data, "lower"))

    def test_nonfinite_dual_does_not_invalidate_a_physical_primal(self):
        with patch.object(core, "audit", return_value={"pass": True, "value": 1.000003, "errors": {}}):
            data = core.evaluate(self.prepared(), np.array([1.000003]), np.array([np.nan, 0.]))
        self.assertTrue(data["physical_pass"])
        self.assertTrue(data["witness"])
        self.assertFalse(data["numerical_gate"])
        self.assertEqual(driver.sign(data, "lower"), 1)

    def test_raw_physical_tolerance_is_not_relaxed_by_normalized_audit(self):
        pr = self.prepared()
        pr.m.stress_scale = 100.
        with patch.object(core, "audit", return_value={"pass": True, "value": 1.000003,
                                                       "errors": {"yield": 5e-9}}):
            data = core.evaluate(pr, np.array([1.000003]))
        self.assertAlmostEqual(data["raw_lower_errors"]["yield"], 5e-7)
        self.assertFalse(data["physical_pass"])
        self.assertFalse(data["witness"])


class ConeScalingTests(unittest.TestCase):
    def test_scalar_and_soc_rows_share_their_block_factor_and_dual_returns_to_raw_units(self):
        G = sparse.csc_matrix([[3., 4.], [2., 0.], [0., 6.], [0., 8.]])
        E = sparse.csc_matrix((0, 2))
        h = np.array([2., 3., 0., 0.])
        c = np.array([5., 4.])
        dims = [1, 3]
        raw = (E, np.array([]), G, h, dims, c)
        norm_soc = np.sqrt(104.)
        restored = []
        pr = SimpleNamespace(original=raw, reduced=raw, cost=c, A=-G, rhs=h,
                cones=[core.clarabel.NonnegativeConeT(1), core.clarabel.SecondOrderConeT(3)],
                objective_offset=0., assembly_seconds=0., transform_seconds=0.,
                m=mesh(), fs=1., mode="upper", q=8, kind="none", restore=lambda x: x.copy())
        def restore_dual(z):
            restored.append(z.copy())
            return z.copy(), {"seconds": 0.}
        pr.restore_dual = restore_dual
        captured = {}
        class Solver:
            def __init__(self, P, cost, A, rhs, cones, options):
                captured.update(A=A.toarray(), rhs=rhs.copy())
            def solve(self):
                return SimpleNamespace(x=np.zeros(2), z=np.array([5., norm_soc, 0., 0.]),
                       status="Solved", iterations=1, solve_time=0., obj_val=0., obj_val_dual=0.)
        with patch.object(core.clarabel, "DefaultSolver", Solver), \
             patch.object(core, "audit", return_value={"pass": True, "value": 0., "errors": {}}):
            data, field, dual = core.solve(pr, cone_scaling=True)
        factors = np.array([1 / 5., 1 / norm_soc, 1 / norm_soc, 1 / norm_soc])
        np.testing.assert_allclose(captured["A"], (-G).toarray() * factors[:, None])
        np.testing.assert_allclose(captured["rhs"], h * factors)
        np.testing.assert_allclose(restored[0], [1., 1., 0., 0.])
        np.testing.assert_allclose(dual, [1., 1., 0., 0.])
        self.assertEqual(data["dual_stationarity"], 0.)
        self.assertTrue(data["physical_pass"])


class SearchBoundaryTests(unittest.TestCase):
    def search(self, callback, initial=(.9, 1.1)):
        attempts = []
        m = mesh()
        def prepared(m, fs, mode, q=8, **kwargs):
            return SimpleNamespace(fs=fs, m=m, mode=mode)
        def solve(pr, *args, **kwargs):
            attempts.append(pr.fs)
            return callback(pr.fs), np.full(12, pr.fs), np.zeros(1)
        with tempfile.TemporaryDirectory(prefix="rpfem-search-test-") as temporary, \
             patch.object(driver, "Prepared", prepared), patch.object(driver, "solve", solve), \
             patch.object(driver.Certificate, "publish", return_value={}), \
             patch.object(driver, "write"), redirect_stdout(io.StringIO()):
            certificate, summary = driver.search(m, "upper", initial, Path(temporary), "synthetic")
        return certificate, summary, attempts

    def test_reversed_or_nonfinite_initial_brackets_are_rejected(self):
        for initial in ((1.1, .9), (1., 1.), (-1., 1.), (np.nan, 1.), (1., np.inf)):
            with self.subTest(initial=initial), self.assertRaises(ValueError):
                self.search(lambda fs: result(value=0.), initial)

    def test_unknown_trials_preserve_the_admissible_endpoint(self):
        def callback(fs):
            if fs == 1.1:
                return result(value=.9)
            if fs == .9:
                return result(value=1.1, witness=False, physical=False,
                              dual_stationarity=0., dual_cone_error=0., dual_objective=1.1)
            return result(value=np.nan, witness=False, physical=False)
        certificate, summary, attempts = self.search(callback)
        self.assertEqual(certificate.fs, 1.1)
        self.assertFalse(certificate.root_complete)
        self.assertFalse(summary["root_complete"])
        self.assertTrue(all(np.isfinite(attempts)))

    def test_nonfinite_primal_value_with_valid_original_dual_uses_bisection(self):
        def callback(fs):
            if fs < 1.:
                return result(value=np.nan, witness=False, physical=False,
                              dual_stationarity=0., dual_cone_error=0., dual_objective=1. + (1. - fs))
            return result(value=.999999 - (fs - 1.), numerical=True)
        certificate, summary, attempts = self.search(callback)
        self.assertIn(1., attempts)
        self.assertTrue(certificate.root_complete)
        self.assertEqual(summary["bracket_signs"], [1, -1])
        self.assertLess(summary["bracket"][0], summary["bracket"][1])
        self.assertLessEqual(summary["bracket_width"], 1e-4 * max(1., sum(summary["bracket"]) / 2))

    def test_all_nonfinite_trials_return_unknown_without_schema_errors(self):
        certificate, summary, attempts = self.search(lambda fs: result(value=np.nan, witness=False, physical=False))
        self.assertIsNone(certificate)
        self.assertFalse(summary["root_complete"])
        self.assertGreater(len(attempts), 2)


class SeedBoundaryTests(unittest.TestCase):
    def run_seed(self, mutate=None, recheck=False, candidates=None, resume=False, field_mode=None, field_q=None):
        lower, upper = mesh("lower"), mesh("upper")
        meta = {"case": "P06A", "history": [], "lower": root_metadata("lower", 1., lower),
                "upper": root_metadata("upper", 1.005, upper)}
        if mutate:
            mutate(meta)
        calls, reports = [], []
        with tempfile.TemporaryDirectory(prefix="rpfem-seed-test-") as temporary:
            folder = Path(temporary)
            seed = folder / "output/P06A" if resume else folder / "seed"
            seed.mkdir(parents=True)
            (seed / "summary.json").write_text(json.dumps(meta), encoding="utf-8")
            for mode, fs in (("lower", 1.), ("upper", 1.005)):
                np.savez_compressed(seed / f"best_{mode}_field.npz", x=np.full(12, fs), fs=fs,
                                    mode=field_mode or mode, q=8 if field_q is None else field_q)
            args = SimpleNamespace(case="P06A", seed=None if resume else seed, resume=resume,
                     recheck=recheck, output=folder / "output", cycles=0, max_elements=100,
                     theta=.6, tolerance=1e-4, solve_seconds=1., fan_lower=False)
            def prepared(m, fs, mode, q=8, **kw):
                return SimpleNamespace(m=m, fs=fs, mode=mode, q=q)
            def audit(pr, x):
                return result(value=1.000003 if pr.mode == "lower" else .999998)
            def publish(certificate, path):
                return {"root_complete": certificate.root_complete, "fs": certificate.fs,
                        "mesh_id": certificate.mesh.mesh_id, "source": certificate.source}
            def search(m, mode, initial, path, label, *other):
                calls.append((mode, m.mesh_id, tuple(initial)))
                fs = 1. if mode == "lower" else 1.005
                candidate_mesh = candidates[mode] if candidates and mode in candidates else m
                certificate = driver.Certificate(mode, fs, candidate_mesh, np.full(12, fs),
                                       result(value=1.000003 if mode == "lower" else .999998), "fresh-root", True)
                return certificate, {"root_complete": True}
            with patch.object(driver, "load_model", return_value=(upper, {})), \
                 patch.object(driver, "load_saved_mesh", side_effect=lambda path, rec: lower if "lower" in path.name else upper), \
                 patch.object(driver, "Prepared", prepared), patch.object(driver, "evaluate", audit), \
                 patch.object(driver.Certificate, "publish", publish), patch.object(driver, "search", search), \
                 patch.object(driver, "save_mesh"), \
                 patch.object(driver, "write", side_effect=lambda path, report: reports.append(deepcopy(report))), \
                 redirect_stdout(io.StringIO()):
                driver.run(args)
        return reports[-1], calls

    def test_matching_complete_seed_and_resume_preserve_target(self):
        for resume in (False, True):
            with self.subTest(resume=resume):
                report, calls = self.run_seed(resume=resume)
                self.assertTrue(report["target_met"])
                self.assertTrue(report["root_complete"])
                self.assertEqual(calls, [])

    def test_every_identity_mismatch_keeps_only_the_physical_witness(self):
        keys = ("case", "mode", "q", "fs", "mesh_id", "physical_model_id")
        for key in keys:
            for place in ("metadata", "root"):
                if key == "fs" and place == "root":
                    continue
                def mutate(meta, key=key, place=place):
                    target = meta["lower"] if place == "metadata" else meta["lower"]["audit"]["discrete_root_search"]
                    target[key] = 4 if key == "q" else 2. if key == "fs" else "DIFFERENT"
                with self.subTest(key=key, place=place):
                    report, calls = self.run_seed(mutate)
                    self.assertFalse(report["target_met"])
                    self.assertFalse(report["lower"]["root_complete"])
                    self.assertEqual(report["lower"]["fs"], 1.)
                    self.assertEqual(calls, [])

    def test_wrong_policy_sign_order_width_endpoint_or_flag_is_not_a_complete_root(self):
        mutations = {
            "policy": lambda r: r.update(policy="DAVIS"),
            "signs": lambda r: r.update(bracket_signs=[-1, 1]),
            "reverse": lambda r: r.update(bracket=[1.00005, 1.]),
            "equal": lambda r: r.update(bracket=[1., 1.]),
            "negative": lambda r: r.update(bracket=[-1., -0.99995]),
            "old_tolerance": lambda r: r.update(bracket=[1., 1.0005], fs_tolerance=.001),
            "endpoint": lambda r: r.update(endpoint=1.00001),
            "endpoint_not_bracket_side": lambda r: r.update(bracket=[.99998, 1.00002]),
            "flag": lambda r: r.update(root_complete=False),
            "nonfinite": lambda r: r.update(bracket=[1., float("inf")]),
            "missing": lambda r: r.clear(),
        }
        for name, change in mutations.items():
            def mutate(meta, change=change):
                change(meta["lower"]["audit"]["discrete_root_search"])
            with self.subTest(name=name):
                report, _ = self.run_seed(mutate)
                self.assertFalse(report["target_met"])
                self.assertFalse(report["lower"]["root_complete"])

    def test_wrong_field_mode_or_quadrature_is_rejected_before_root_use(self):
        with self.assertRaisesRegex(ValueError, "mode or quadrature"):
            self.run_seed(field_q=4)
        with self.assertRaisesRegex(ValueError, "mode or quadrature"):
            self.run_seed(field_mode="upper")

    def test_same_fs_recheck_upgrades_root_on_each_original_witness_mesh(self):
        def mutate(meta):
            for mode in ("lower", "upper"):
                meta[mode]["root_complete"] = False
        report, calls = self.run_seed(mutate, recheck=True)
        self.assertTrue(report["target_met"])
        self.assertTrue(report["lower"]["root_complete"])
        self.assertTrue(report["upper"]["root_complete"])
        self.assertEqual([(mode, identity) for mode, identity, bracket in calls],
                         [("lower", "mesh-lower"), ("upper", "mesh-upper")])
        self.assertEqual(report["lower"]["source"], "fresh-root")

    def test_candidate_from_another_physical_model_is_rejected(self):
        def mutate(meta):
            meta["lower"]["root_complete"] = False
        with self.assertRaisesRegex(ValueError, "Physical models differ"):
            self.run_seed(mutate, recheck=True, candidates={"lower": mesh("lower", "physical-B")})


if __name__ == "__main__":
    unittest.main(verbosity=2)
