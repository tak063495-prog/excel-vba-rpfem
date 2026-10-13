"""Auxiliary QP fields must pass the original physical gates independently."""
from pathlib import Path
from types import SimpleNamespace
import json
import sys
import tempfile
import unittest

sys.path.insert(0, str(Path(__file__).resolve().parents[1] / 'python'))
import accuracy_driver as driver
import c0_slope_wedge_driver as wedge
import experiment_core as core
import upper_proximal_candidate as candidate
from test_inverse_dual_guide import physical_triangle


class UpperProximalCandidateTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        _, record = driver.load_model('D02A')
        cls.mesh, _, _ = wedge.wedge_mesh(record, .01)

    def source(self, folder, fs=1.6):
        prepared = core.Prepared(self.mesh, fs, 'upper', 8)
        driver.save_mesh(folder / 'original_mesh.npz', self.mesh)
        core.np.savez_compressed(folder / 'original_candidate.npz',
                                x=core.np.zeros(len(prepared.original[-1])),
                                fs=fs, mode='upper', q=8)
        # Deliberately false legacy claims: these are never imported.
        core.write(folder / 'audit.json', {
            'case': 'D02A', 'fs': fs, 'q': 8, 'mesh_id': self.mesh.mesh_id,
            'physical_model_id': self.mesh.physical_model_id,
            'root_complete': True, 'physical_pass': True, 'status': 'Solved'})

    def test_recovery_from_zero_work_center_preserves_original_and_excludes_modified_dual(self):
        prepared = core.Prepared(self.mesh, 1.6, 'upper', 8)
        source = core.np.zeros(len(prepared.original[-1]))
        self.assertFalse(core.evaluate(prepared, source)['physical_pass'])
        before = candidate.program_hash(prepared.original)
        field, dual, report = candidate.solve_candidate(prepared, source, seconds=10.)
        audit = core.evaluate(prepared, field)
        self.assertTrue(audit['physical_pass']); self.assertTrue(audit['witness'])
        self.assertEqual(before, candidate.program_hash(prepared.original))
        self.assertEqual(len(dual), len(prepared.original[1]) + len(prepared.original[3]) + 1)
        self.assertFalse(report['root_complete'])
        self.assertFalse(report['modified_dual_is_original_root_guide'])
        self.assertNotIn('dual_stationarity', report['original_audit'])
        self.assertFalse(core.inverse_dual_guide(prepared, report['original_audit'])['verified'])
        core.np.testing.assert_array_equal(source, 0.)

    def test_infeasible_physical_strength_cannot_publish_from_solver_status(self):
        prepared = core.Prepared(self.mesh, 1.3, 'upper', 8)
        field, _, report = candidate.solve_candidate(
            prepared, core.np.zeros(len(prepared.original[-1])), seconds=10.)
        self.assertFalse(core.evaluate(prepared, field)['witness'])
        self.assertFalse(report['root_complete'])

    def test_cohesive_original_objective_cap_accepts_and_rejects_actual_fields(self):
        for fs, expected, scaling in [(100., True, False), (.01, False, False),
                                       (100., True, True), (.01, False, True)]:
            prepared = core.Prepared(physical_triangle(), fs, 'upper', 8)
            self.assertTrue(core.np.any(prepared.original[-1] != 0.))
            field, _, report = candidate.solve_candidate(
                prepared, core.np.zeros(len(prepared.original[-1])), seconds=10.,
                cone_scaling=scaling)
            original = core.evaluate(prepared, field)
            self.assertEqual(original['physical_pass'] and original['witness'], expected)
            if expected:
                self.assertLessEqual(original['objective'], report['cap'] + core.PHYSICAL_TOL)
            self.assertFalse(core.inverse_dual_guide(prepared, report['original_audit'])['verified'])

    def test_cli_ignores_source_claims_copies_exact_sources_and_keeps_candidate_root_incomplete(self):
        with tempfile.TemporaryDirectory() as directory:
            root = Path(directory); source = root / 'seed'; source.mkdir()
            self.source(source)
            hashes = {p.name: candidate.file_hash(p) for p in source.iterdir()}
            output = root / 'output'
            report = candidate.run(SimpleNamespace(case='D02A', seed=source,
                output=output, fs=None, alpha=.01, seconds=10., cap=core.UPPER_TARGET - 1e-6))
            self.assertTrue(report['original_audit']['witness'])
            self.assertFalse(report['source_provenance']['reference_original_audit']['physical_pass'])
            certificate = json.loads((output / 'best_upper_certificate.json').read_text())
            self.assertFalse(certificate['root_complete'])
            self.assertNotIn('discrete_root_search', certificate['audit'])
            self.assertEqual(hashes, {p.name: candidate.file_hash(p) for p in source.iterdir()})
            self.assertEqual(hashes, {p.name: candidate.file_hash(p) for p in (output / 'source_files').iterdir()})
            self.assertEqual(report['candidate_code_sha256'], candidate.file_hash(output / 'experiment_source.py'))
            self.assertEqual(report['implementation_sha256'], {
                name: candidate.file_hash(output / 'implementation_sources' / name)
                for name in report['implementation_sha256']})
            with core.np.load(output / 'original_candidate.npz', allow_pickle=False) as saved:
                self.assertIn('modified_dual', saved.files); self.assertNotIn('z', saved.files)

    def test_original_identity_and_unchanged_gate_parameters_are_required(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory); self.source(folder)
            path = folder / 'audit.json'; metadata = json.loads(path.read_text())
            for key, value in [('case', 'D03N'), ('q', 4), ('fs', 1.5),
                               ('mesh_id', 'different'), ('physical_model_id', 'different')]:
                core.write(path, dict(metadata, **{key: value}))
                with self.assertRaises(ValueError): candidate.load_center('D02A', folder)
            core.write(path, metadata)
        prepared = core.Prepared(self.mesh, 1.6, 'upper', 8)
        source = core.np.zeros(len(prepared.original[-1]))
        for options in [dict(alpha=0.), dict(seconds=0.), dict(alpha=core.np.nan),
                        dict(cap=core.UPPER_TARGET), dict(cap=core.np.inf)]:
            with self.assertRaises(ValueError): candidate.solve_candidate(prepared, source, **options)
        source[0] = core.np.nan
        with self.assertRaises(ValueError): candidate.solve_candidate(prepared, source)
        prepared = core.Prepared(physical_triangle(), 100., 'upper', 8)
        prepared.original[-1][:] = 0.
        with self.assertRaises(ValueError):
            candidate.solve_candidate(prepared, core.np.zeros(len(prepared.original[-1])))


if __name__ == '__main__': unittest.main()
