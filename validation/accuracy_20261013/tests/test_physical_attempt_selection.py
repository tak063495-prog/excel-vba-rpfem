"""Recovery keeps the best admissible iterate without promoting rejected fields."""
from pathlib import Path
import sys
import unittest

sys.path.insert(0,str(Path(__file__).resolve().parents[1]/'python'))
import accuracy_driver as driver
import experiment_core as core

EVIDENCE=driver.REPO/'validation/accuracy_20261013/snapshots/20261013_n_nested_candidate_audit'


class PhysicalAttemptSelectionTests(unittest.TestCase):
    def source(self,mode,entry):
        _,rec=driver.load_model('P06N');folder=EVIDENCE/entry
        mesh=driver.load_saved_mesh(folder/f'best_{mode}_mesh.npz',rec)
        with core.np.load(folder/f'best_{mode}_field.npz',allow_pickle=False) as data:
            x,fs=data['x'].copy(),float(data['fs'])
        return core.Prepared(mesh,fs,mode,8,reduction='none'),x

    def test_later_real_admissible_lower_iterate_is_not_hidden_by_first_iterate(self):
        prepared,source=self.source('lower','parent_lower5422')
        core.np.testing.assert_array_equal(prepared.original[1],0.)
        # Scaling gravity-only free-load equilibrium preserves admissibility.
        # Both loads miss the witness gate; the closer field enables the
        # driver's existing fixed-load recovery trigger at lambda >= .98.
        first=core.evaluate(prepared,source*(.978/source[-1]))
        later=core.evaluate(prepared,source*(.998/source[-1]))
        self.assertTrue(first['physical_pass']);self.assertTrue(later['physical_pass'])
        self.assertFalse(first['witness']);self.assertFalse(later['witness'])
        selected=max([first,later],key=lambda r:driver.physical_attempt_rank(r,'lower'))
        self.assertIs(selected,later)
        self.assertIsNone(driver.sign(selected,'lower'))
        self.assertFalse(selected['witness'])

    def test_rejected_real_upper_velocity_cannot_displace_a_witness(self):
        prepared,source=self.source('upper','parent_upper6106')
        original=source.copy();valid=core.evaluate(prepared,source)
        corrupted=source.copy();corrupted[0]+=1.
        rejected=core.evaluate(prepared,corrupted)
        self.assertTrue(valid['physical_pass']);self.assertTrue(valid['witness'])
        self.assertFalse(rejected['physical_pass']);self.assertFalse(rejected['witness'])
        rejected['value']=-1e100
        selected=max([valid,rejected],key=lambda r:driver.physical_attempt_rank(r,'upper'))
        self.assertIs(selected,valid)
        core.np.testing.assert_array_equal(source,original)


if __name__=='__main__':unittest.main()
