"""Keep the three-fix release limited to the four declared VBA modules."""
from pathlib import Path
import hashlib
import unittest

ROOT=Path(__file__).resolve().parents[1]
BASE=ROOT.parent/'G1R20'
CHANGED={'RPX_UI.bas','RPX_Water.bas','RPX_Seep.bas','RPX_SeepUI.bas'}

class ReleaseScope(unittest.TestCase):
    def test_full_source_and_exactly_four_changes(self):
        sources={p.name:p for p in (ROOT/'src/vba').glob('*.bas')}
        self.assertEqual(len(sources),63)
        self.assertEqual(set(sources),{p.name for p in (BASE/'src/vba').glob('*.bas')})
        changed={n for n,p in sources.items() if p.read_bytes()!=(BASE/'src/vba'/n).read_bytes()}
        self.assertEqual(changed,CHANGED)

    def test_import_text_and_crlf(self):
        for p in (ROOT/'src/vba').glob('*.bas'):
            with self.subTest(module=p.name):
                raw=p.read_bytes()
                self.assertEqual(raw.decode('utf8'),(ROOT/'import'/p.name).read_bytes().decode('cp932'))
                self.assertNotIn(b'\n',raw.replace(b'\r\n',b''))

    def test_original_numerical_kernels_unchanged(self):
        for name in ['RPX_AccuracyNative.bas','RPX_Adapt.bas','RPX_Fs.bas','RPX_Refine.bas',
                     'RPX_FlowMesh.bas','RPX_Audit.bas','RPX_Mesh.bas','RPX_Solve.bas']:
            self.assertEqual((ROOT/'src/vba'/name).read_bytes(),(BASE/'src/vba'/name).read_bytes())

if __name__=='__main__':unittest.main()
