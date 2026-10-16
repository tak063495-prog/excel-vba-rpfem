"""The Fix2 review changes lifecycle handling and preserves numerical kernels."""
from pathlib import Path
import unittest

ROOT = Path(__file__).resolve().parents[1]
BASE = ROOT.parent / 'G1R20_Fix1'
CHANGED = {'RPX_UI.bas', 'RPX_Control.bas', 'RPX_Water.bas'}


class ReleaseScope(unittest.TestCase):
    def test_only_declared_modules_changed(self):
        sources = {p.name: p for p in (ROOT / 'src/vba').glob('*.bas')}
        self.assertEqual(len(sources), 63)
        self.assertEqual(set(sources), {p.name for p in (BASE / 'src/vba').glob('*.bas')})
        changed = {name for name, path in sources.items()
                   if path.read_bytes() != (BASE / 'src/vba' / name).read_bytes()}
        self.assertEqual(changed, CHANGED)

    def test_import_matches_source_and_crlf(self):
        for path in (ROOT / 'src/vba').glob('*.bas'):
            with self.subTest(module=path.name):
                raw = path.read_bytes()
                self.assertEqual(raw.decode('utf8'), (ROOT / 'import' / path.name).read_bytes().decode('cp932'))
                self.assertNotIn(b'\n', raw.replace(b'\r\n', b''))

    def test_numerical_kernels_preserved(self):
        for name in ['RPX_AccuracyNative.bas', 'RPX_Adapt.bas', 'RPX_Fs.bas',
                     'RPX_Refine.bas', 'RPX_FlowMesh.bas', 'RPX_Audit.bas',
                     'RPX_Mesh.bas', 'RPX_Solve.bas', 'RPX_Seep.bas', 'RPX_SeepUI.bas']:
            self.assertEqual((ROOT / 'src/vba' / name).read_bytes(), (BASE / 'src/vba' / name).read_bytes())


if __name__ == '__main__':
    unittest.main()
