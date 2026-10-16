"""Small rejection tests; no real artifacts, native calls, or solver calls."""
import collections, struct, unittest
import audit_cold_native as audit


def exact(value):
    return struct.pack('<d', value).hex().upper()


class AuditContractTests(unittest.TestCase):
    def meta(self, low=1.5, high=1.5001220703125):
        return {key:[value] for key,value in dict(lower_root_valid='True', search_incomplete='False',
                    lower_root_a_exact=exact(low),lower_root_b_exact=exact(high)).items()}

    def test_original_relative_root_contract(self):
        result = audit.root_contract(self.meta(), 'lower', 1.5)
        self.assertGreater(result['width_absolute'], 1e-4)
        self.assertLess(result['width_normalized'], 1e-4)

    def test_wrong_accepted_endpoint_rejected(self):
        with self.assertRaisesRegex(ValueError, 'own accepted root endpoint'):
            audit.root_contract(self.meta(), 'lower', 1.50001)

    def test_incomplete_or_excessive_root_rejected(self):
        for meta in (self.meta(high=1.5002), dict(self.meta(),search_incomplete=['True'])):
            with self.assertRaises(ValueError):
                audit.root_contract(meta, 'lower', 1.5)

    def test_nonfinite_and_truncated_proof_rejected(self):
        with self.assertRaisesRegex(ValueError, 'Nonfinite'):
            audit.exact(exact(float('nan')))
        with self.assertRaisesRegex(ValueError, 'Truncated'):
            audit.state(b'RP25M001'+struct.pack('<5i',3,1,1,0,8))

    def test_foreign_magic_and_extra_metadata_rejected(self):
        with self.assertRaisesRegex(ValueError, 'Unsupported binary magic'):
            audit.Reader(b'AC18P999', b'AC18P001')
        with self.assertRaisesRegex(ValueError, 'Unexpected trailing'):
            audit.Reader(b'AC18P001x', b'AC18P001').end()


if __name__ == '__main__':
    unittest.main()
