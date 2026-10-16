"""Source-level contract tests; no Excel, numerical solver, or output mutation."""
from pathlib import Path
import importlib.util
import json
import tempfile
import unittest

ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location('state_reset_manifest', ROOT/'make_overlay_manifest.py')
overlay = importlib.util.module_from_spec(spec)
spec.loader.exec_module(overlay)


class ContractTests(unittest.TestCase):
    def test_exactly_four_module_pairs(self):
        expected = set(overlay.EDITS)
        self.assertEqual({p.name for p in (ROOT/'src/vba').glob('*.bas')}, expected)
        self.assertEqual({p.name for p in (ROOT/'import').glob('*.bas')}, expected)

    def test_one_call_insertion_per_module_no_other_text_change(self):
        for name, (relative, before, after) in overlay.EDITS.items():
            with self.subTest(module=name):
                base = overlay.normalize((overlay.REPO/relative).read_text(encoding='utf8')).splitlines()
                current = overlay.normalize((ROOT/'src/vba'/name).read_text(encoding='utf8')).splitlines()
                self.assertEqual(len(base), len(current))
                changes = [(a,b) for a,b in zip(base,current) if a != b]
                self.assertEqual(changes, [(before, after)])
                self.assertEqual(after.replace('Call ', '', 1), before)

    def test_correct_baselines(self):
        self.assertTrue(overlay.EDITS['RPX_Fs.bas'][0].startswith('versions/G1R19_2pct/'))
        for name in ['RPX_Conic.bas','RPX_Solve.bas','RPX_Seep.bas']:
            self.assertTrue(overlay.EDITS[name][0].startswith('versions/G1R19/'))

    def test_manifest_and_all_reference_bytes(self):
        saved = json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf8'))
        self.assertEqual(saved, overlay.manifest_data())
        self.assertEqual(saved['excel_runs'], 0)
        self.assertEqual(saved['solver_calls'], 0)
        self.assertFalse(saved['production_artifacts_modified'])

    def test_import_decodes_identically_and_crlf_preserved(self):
        for name in overlay.EDITS:
            with self.subTest(module=name):
                raw = (ROOT/'src/vba'/name).read_bytes()
                self.assertEqual(raw.decode('utf8'), (ROOT/'import'/name).read_bytes().decode('cp932'))
                self.assertNotIn(b'\n', raw.replace(b'\r\n', b''))

    def test_gate_and_root_text_is_preserved_by_reverse_patch(self):
        # Reversing only the explicit-call edit must recover the entire baseline,
        # not merely a sample of gate constants or selected physical formulas.
        for name, (relative,before,after) in overlay.EDITS.items():
            with self.subTest(module=name):
                candidate = overlay.normalize((ROOT/'src/vba'/name).read_text(encoding='utf8'))
                self.assertEqual(candidate.replace(after+'\n',before+'\n'),
                                 overlay.normalize((overlay.REPO/relative).read_text(encoding='utf8')))

    def test_product_sources_have_no_test_traps(self):
        for path in (ROOT/'src/vba').glob('*.bas'):
            self.assertNotIn('SRProbe', path.read_text(encoding='utf8'))

    def test_probe_generator_refuses_existing_directory(self):
        module_spec = importlib.util.spec_from_file_location('state_reset_probes', ROOT/'tests/make_probe_sources.py')
        probes = importlib.util.module_from_spec(module_spec)
        module_spec.loader.exec_module(probes)
        with tempfile.TemporaryDirectory() as directory:
            with self.assertRaises(FileExistsError):
                probes.generate(Path(directory))

    def test_probe_sources_are_separate_and_marked_test_only(self):
        module_spec = importlib.util.spec_from_file_location('state_reset_probes2', ROOT/'tests/make_probe_sources.py')
        probes = importlib.util.module_from_spec(module_spec)
        module_spec.loader.exec_module(probes)
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory)/'probe'
            result = probes.generate(target)
            self.assertTrue(result['test_only'])
            self.assertFalse(result['numerical_evidence_allowed'])
            self.assertEqual(result['modules'], 63)
            conic = (target/'src/vba/RPX_Conic.bas').read_text(encoding='utf8')
            self.assertIn('Call SRProbeCompileTrap', conic)
            self.assertIn('SRProbeConicEntry', conic)
            solve = (target/'src/vba/RPX_Solve.bas').read_text(encoding='utf8')
            self.assertIn('SRProbePrimeSolveSentinels', solve)
            self.assertIn('SRProbeKrylovEntry', solve)
            seep = (target/'src/vba/RPX_Seep.bas').read_text(encoding='utf8')
            self.assertIn('Call SRProbeSeepTrap', seep)

    def test_negative_control_retains_exact_original_four_lines(self):
        module_spec = importlib.util.spec_from_file_location('state_reset_negative_probes', ROOT/'tests/make_probe_sources.py')
        probes = importlib.util.module_from_spec(module_spec)
        module_spec.loader.exec_module(probes)
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory)/'negative'
            result = probes.generate(target, negative_control=True)
            self.assertTrue(result['negative_control'])
            for name, (_,before,after) in overlay.EDITS.items():
                current = (target/'src/vba'/name).read_text(encoding='utf8')
                expected=(before.replace(': HEnsureDD',': Call SRProbeKrylovAfterEntryTrap: HEnsureDD')
                          if name=='RPX_Solve.bas' else before)
                self.assertIn(expected+'\n', current)
                self.assertNotIn(after+'\n', current)
            solve = (target/'src/vba/RPX_Solve.bas').read_text(encoding='utf8')
            self.assertEqual(solve.count('Call SRProbeKrylovAfterEntryTrap:'),1)


if __name__ == '__main__':
    unittest.main()
