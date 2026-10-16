"""A running checkpoint or failed process must not become a release PASS."""
from pathlib import Path
import hashlib
import json
import tempfile
import unittest
from unittest.mock import patch

import collect_current as collector


class SupervisorEvidenceTests(unittest.TestCase):
    def setUp(self):
        self.temp = tempfile.TemporaryDirectory()
        self.repo = Path(self.temp.name).resolve()
        self.batch = self.repo/'variant'/'results'/'cold'
        self.folder = self.batch/'cases'/'P06A'
        self.folder.mkdir(parents=True)
        self.source = self.repo/'numeric.py'
        self.source.write_text('original numerical source\n', encoding='utf8')
        digest = hashlib.sha256(self.source.read_bytes()).hexdigest()
        self.owner = dict(status='FINISHED', source_unchanged=True,
                          source_sha256={str(self.source): digest},
                          source_sha256_after={str(self.source): digest},
                          cases={'P06A': dict(folder=str(self.folder), phase='FINISHED',
                                             returncode=0, audit_returncode=0,
                                             independent_audit=str(self.batch/'audit.json'),
                                             started_epoch=10, finished_epoch=30)})
        self.summary = dict(case='P06A', status='TARGET_MET', target_met=True,
                            lower={'fs': .84}, upper={'fs': .85},
                            relative_width_symmetric=2*(.85-.84)/(.85+.84))

    def tearDown(self):
        self.temp.cleanup()

    def collect(self):
        (self.batch/'OWNERSHIP.json').write_text(json.dumps(self.owner), encoding='utf8')
        (self.folder/'summary.json').write_text(json.dumps(self.summary), encoding='utf8')
        with patch.object(collector, 'REPO', self.repo), patch.object(collector, 'validate', return_value={}) as audit:
            result = list(collector.supervised_attempts(self.batch))[0]
        return result, audit

    def test_complete_run_requires_independent_audit(self):
        result, audit = self.collect()
        self.assertTrue(result['verified'])
        audit.assert_called_once()
        self.assertEqual(result['seconds'], 20)

    def test_running_checkpoint_is_never_a_pass(self):
        self.owner['status'] = 'RUNNING'
        result, audit = self.collect()
        self.assertFalse(result['verified'])
        audit.assert_not_called()

    def test_interrupted_comparison_does_not_remain_running(self):
        self.owner['cases']['P06A']['returncode'] = 15
        self.summary.update(status='RUNNING', target_met=False)
        result, audit = self.collect()
        self.assertEqual(result['status'], 'INTERRUPTED')
        self.assertFalse(result['verified'])
        audit.assert_not_called()

    def test_continuation_is_separate_from_a_completed_cold_run(self):
        self.owner['cases']['P06A']['command'] = ['python', 'run_case.py', '--prior-run', 'fresh-v1']
        self.summary.update(cold_cumulative_seconds=926, adopted_path_seconds=926,
                            all_current_session_trial_seconds=1513)
        result, audit = self.collect()
        self.assertTrue(result['verified'])
        self.assertFalse(result['cold_execution_verified'])
        self.assertEqual(result['kind'], 'continuation')
        self.assertEqual(result['seconds'], 926)
        self.assertEqual(result['process_seconds'], 20)
        self.assertEqual(result['all_current_session_trial_seconds'], 1513)
        audit.assert_called_once()

    def test_unidentified_saved_seed_is_not_counted_as_fresh(self):
        self.summary['execution_scope'] = {'saved_witness_seed_requested': True}
        result, audit = self.collect()
        self.assertEqual(result['status'], 'EVIDENCE_REJECTED')
        self.assertFalse(result['verified'])
        audit.assert_not_called()

    def test_failed_worker_or_auditor_is_rejected(self):
        for key in ('returncode', 'audit_returncode'):
            with self.subTest(key=key):
                self.owner['cases']['P06A'][key] = 1
                result, audit = self.collect()
                self.assertEqual(result['status'], 'EVIDENCE_REJECTED')
                self.assertFalse(result['verified'])
                audit.assert_not_called()
                self.owner['cases']['P06A'][key] = 0

    def test_changed_source_cannot_reuse_finished_evidence(self):
        self.source.write_text('changed numerical source\n', encoding='utf8')
        result, audit = self.collect()
        self.assertEqual(result['status'], 'EVIDENCE_REJECTED')
        audit.assert_not_called()

    def test_incomplete_source_freeze_is_rejected(self):
        self.owner['source_sha256_after'] = {}
        result, audit = self.collect()
        self.assertEqual(result['status'], 'EVIDENCE_REJECTED')
        audit.assert_not_called()


if __name__ == '__main__':
    unittest.main()
