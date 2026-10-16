"""Reject falsely relabelled or changed component evidence without mutating it."""
import copy
from pathlib import Path
import unittest
from unittest.mock import patch

import composite_provenance as proof


PAIR = proof.REPO/'validation/G1R19_python26_geometry_replay_20261016/results/P06A_COMPOSITE_20261016T122851Z'
NPAIR = proof.REPO/'validation/G1R19_python26_dual_linear_repair_20261016/results/P06N_COMPOSITE_20261016T125240Z'


@unittest.skipUnless((PAIR/'summary.json').exists(), 'Local component evidence fixture absent')
class CompositeEvidenceTests(unittest.TestCase):
    def run_proof(self, transform=None):
        original = proof.read
        def changed(path):
            value = original(path)
            return transform(Path(path), copy.deepcopy(value)) if transform else value
        audit = original(PAIR.with_name(PAIR.name+'_INDEPENDENT_AUDIT.json'))
        with patch.object(proof, 'read', side_effect=changed):
            return proof.verify_composite_provenance(PAIR/'summary.json', 'P06A', audit=audit)

    def test_real_components_pass_without_cold_claim(self):
        result = self.run_proof()
        self.assertTrue(result['source_provenance_verified'])
        self.assertFalse(result['cold_execution_status_verified'])
        self.assertTrue(result['posthoc_composite'])
        self.assertEqual(len(result['components']), 6)

    def test_cold_relabel_rejected(self):
        def change(path, value):
            if path == PAIR/'summary.json':
                value['execution_scope']['generic_cold_start'] = True
            return value
        with self.assertRaisesRegex(ValueError, 'cold run'):
            self.run_proof(change)

    def test_historical_field_relabel_rejected(self):
        def change(path, value):
            if path == PAIR/'summary.json':
                value['execution_scope']['historical_field_imported'] = True
            return value
        with self.assertRaisesRegex(ValueError, 'scope'):
            self.run_proof(change)

    def test_component_end_map_mismatch_rejected(self):
        def change(path, value):
            if path.name == 'COMPONENT_MANIFEST.json':
                entry = value['components'][0]['component_end_source_sha256']
                entry[next(iter(entry))] = '0'*64
            return value
        with self.assertRaisesRegex(ValueError, 'start/end'):
            self.run_proof(change)

    def test_actual_supervisor_exit_mismatch_rejected(self):
        def change(path, value):
            if path.name == 'OWNERSHIP.json':
                for job in value['jobs'].values():
                    job['returncode'] = 99
            return value
        with self.assertRaisesRegex(ValueError, 'exit/audit'):
            self.run_proof(change)

    def test_original_root_context_mismatch_rejected(self):
        def change(path, value):
            if path == PAIR/'best_lower_certificate.json':
                value['audit']['discrete_root_search']['physical_model_id'] = '0'*64
            return value
        with self.assertRaisesRegex(ValueError, 'Root case'):
            self.run_proof(change)


@unittest.skipUnless((NPAIR/'summary.json').exists(), 'Local N component evidence fixture absent')
class UnownedRepairEvidenceTests(unittest.TestCase):
    def run_proof(self, transform=None):
        original = proof.read
        def changed(path):
            value = original(path)
            return transform(Path(path), copy.deepcopy(value)) if transform else value
        audit = original(NPAIR.with_name(NPAIR.name+'_INDEPENDENT_AUDIT.json'))
        with patch.object(proof, 'read', side_effect=changed):
            return proof.verify_composite_provenance(NPAIR/'summary.json', 'P06N', audit=audit)

    def test_real_unowned_repair_does_not_become_cold_run(self):
        result = self.run_proof()
        self.assertFalse(result['cold_execution_status_verified'])
        self.assertTrue(result['independent_original_dual_audits'])
        self.assertEqual(result['timing']['unsupervised_repair_process_wall_seconds'], None)

    def test_invented_process_timing_rejected(self):
        def change(path, value):
            if path.name == 'COMPONENT_MANIFEST.json':
                for component in value['components']:
                    if component.get('ownership_file') is None:
                        component['started_epoch'] = 1234.
            return value
        with self.assertRaisesRegex(ValueError, 'invent process timing'):
            self.run_proof(change)

    def test_optimizer_scope_relabel_rejected(self):
        def change(path, value):
            if path.name == 'summary.json' and path.parent.name == 'P06N_repair_20261016T125042Z':
                value['execution_scope']['optimizer_calls'] = 1
            return value
        with self.assertRaisesRegex(ValueError, 'no-optimizer continuation'):
            self.run_proof(change)

    def test_false_independent_dual_residual_rejected(self):
        def change(path, value):
            if path.name == 'P06N_repair_20261016T125042Z_INDEPENDENT_DUAL_AUDIT.json':
                value['audit']['dual_stationarity'] = 2e-7
            return value
        with self.assertRaisesRegex(ValueError, 'original dual raw gate'):
            self.run_proof(change)

    def test_extra_wrapper_time_cannot_be_called_process_time(self):
        def change(path, value):
            if path.name == 'COMPONENT_MANIFEST.json':
                value['timing']['all_recipe_trial_process_sum_seconds'] += value['timing']['unsupervised_repair_wrapper_seconds']
            return value
        with self.assertRaisesRegex(ValueError, 'process-duration sum'):
            self.run_proof(change)


if __name__ == '__main__':
    unittest.main()
