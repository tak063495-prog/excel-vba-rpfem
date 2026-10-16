"""Portable integration contract for the exact phi-zero Fs solve path.

The Excel regression exercises the solver. These checks protect argument flow,
the default direct-DLL route, the minimal numerical source change, and its oracle.
"""
from pathlib import Path
import csv
import math
import re
import unittest

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT.parent / "G1R20_Fix2"


def body(source, name):
    match = re.search(r"^(?:Public|Private) Function " + name + r"\b.*?^End Function$",
                      source, re.M | re.S)
    if not match:
        raise AssertionError("Missing function: " + name)
    return match.group()


class ExactFsContract(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.source = (ROOT / "src/vba/RPX_Fs.bas").read_text(encoding="utf8")

    def test_exact_request_reaches_bound_without_changing_default_route(self):
        endpoint = body(self.source, "RPX_FsEndpoint")
        exact = endpoint.split("If allZeroPhi Then", 1)[1].split("meshKey = FsMeshKey()", 1)[0]
        self.assertIn('FsValue(1#, mode, target, False)', exact)
        self.assertIn('If XStatus <> "OPTIMAL"', exact)
        self.assertIn('RPX_Bound(mode, factor, allowRecovery)', body(self.source, "FsValue"))
        bound = body(self.source, "RPX_Bound")
        self.assertIn('Optional ByVal allowDirect As Boolean = True', bound.splitlines()[0])
        self.assertIn('If allowDirect And RPX_DirectEnabled() Then', bound)
        self.assertNotIn('RPX_SetSetting', exact + bound)

    def test_only_argument_flow_changed_in_the_entire_fs_module(self):
        baseline = (BASE / "src/vba/RPX_Fs.bas").read_text(encoding="utf8")
        replacements = [
            ('Optional ByVal strengthFactor As Double = 1#) As Double',
             'Optional ByVal strengthFactor As Double = 1#, Optional ByVal allowDirect As Boolean = True) As Double'),
            ('If RPX_DirectEnabled() Then', 'If allowDirect And RPX_DirectEnabled() Then'),
            ('FsValue = RPX_Bound(mode, factor) - target',
             'FsValue = RPX_Bound(mode, factor, allowRecovery) - target'),
        ]
        for old, new in replacements:
            self.assertEqual(baseline.count(old), 1)
            baseline = baseline.replace(old, new)
        self.assertEqual(self.source, baseline)
        self.assertEqual(self.source, (ROOT / "import/RPX_Fs.bas").read_text(encoding="cp932"))

    def test_shear_fixture_has_independent_matching_primal_and_dual_oracles(self):
        for cohesion in (10, 11):
            with (ROOT / f"tests/read_model_shear_c{cohesion}.txt").open(newline="") as stream:
                rows = list(csv.reader(stream))
            self.assertEqual(list(map(int, rows[0])), [4, 2, 1, 0, 4, 1, 1, 8])
            self.assertEqual(list(map(float, rows[1])), [float(cohesion), 0., 0.])
            xy = [tuple(map(float, row)) for row in rows[2:6]]
            triangles = [list(map(int, row[:3])) for row in rows[6:8]]
            area = sum(((xy[b][0]-xy[a][0])*(xy[c][1]-xy[a][1]) -
                        (xy[c][0]-xy[a][0])*(xy[b][1]-xy[a][1]))/2 for a,b,c in triangles)
            self.assertEqual(area, 1.)
            # vx=y, vy=0 is incompressible, fixed on y=0, with unit shear.
            # Constant tau=cohesion saturates the phi=0 yield cone.
            self.assertEqual(math.sqrt(4.*cohesion*cohesion), 2.*cohesion)
            work = 0.
            for row in rows[8:]:
                a,b = map(int,row[:2]); kind=row[2]
                if kind == 'fixed':
                    self.assertEqual((xy[a][1],xy[b][1]), (0.,0.))
                    continue
                dx,dy = xy[b][0]-xy[a][0],xy[b][1]-xy[a][1]
                length=math.hypot(dx,dy); nx,ny=dy/length,-dx/length
                tx,ty=map(float,row[6:8])
                self.assertAlmostEqual(tx,ny)
                self.assertAlmostEqual(ty,nx)
                work += length*tx*(xy[a][1]+xy[b][1])/2
            self.assertEqual(work, 1.)
            self.assertEqual(cohesion*area/work, cohesion)


if __name__ == "__main__":
    unittest.main()
