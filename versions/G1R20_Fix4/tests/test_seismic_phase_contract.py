"""The real Excel history regression supplies the numerical oracle.

Protect the small phase-order fix and make unexpected mathematical edits visible.
"""
from pathlib import Path
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT.parent / "G1R20_Fix3"


class SeismicPhaseContract(unittest.TestCase):
    def test_only_new_raw_phase_precedes_fresh_load_assembly(self):
        source = (ROOT / "src/vba/RPX_Mesh.bas").read_text(encoding="utf8")
        baseline = (BASE / "src/vba/RPX_Mesh.bas").read_text(encoding="utf8")
        added = (
            "    ' Rebuild raw loads before seismic assembly; a previous Fs solve leaves reference_total.\n"
            '    RPX_LoadPhase = "raw"\n'
        )
        self.assertEqual(source.count(added), 1)
        self.assertEqual(source.replace(added, "", 1), baseline)
        procedure = source.split("Public Sub RPX_AssignRegions()", 1)[1].split("End Sub", 1)[0]
        self.assertLess(procedure.index(added), procedure.index("RPX_SeismicBody count"))
        self.assertLess(procedure.rindex('RPX_LoadPhase = "raw"'), procedure.index("Call RPX_WaterApply"))
        self.assertEqual(source, (ROOT / "import/RPX_Mesh.bas").read_text(encoding="cp932"))

    def test_seismic_equation_and_original_bounds_are_unchanged(self):
        for name in ["RPX_Fs.bas", "RPX_Audit.bas", "RPX_Bearing.bas"]:
            self.assertEqual((ROOT / "src/vba" / name).read_bytes(),
                             (BASE / "src/vba" / name).read_bytes(), name)
        seismic=(ROOT / "src/vba/RPX_Seismic.bas").read_text(encoding="utf8")
        old=(BASE / "src/vba/RPX_Seismic.bas").read_text(encoding="utf8")
        for procedure in ["RPX_SeismicBody", "RPX_SeismicKhSigned"]:
            pattern=r"^Public (?:Sub|Function) "+procedure+r"\b.*?^End (?:Sub|Function)$"
            self.assertEqual(re.search(pattern,seismic,re.M|re.S).group(),
                             re.search(pattern,old,re.M|re.S).group(),procedure)


if __name__ == "__main__":
    unittest.main()
