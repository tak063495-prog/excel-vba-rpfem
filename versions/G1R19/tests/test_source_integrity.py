"""Release-level checks for source provenance and the one-DLL distribution."""
from pathlib import Path
import hashlib
import json
import unittest
import zipfile

ROOT = Path(__file__).resolve().parents[1]
VERSIONS = ROOT.parent


class ReleaseIntegrityTests(unittest.TestCase):
    def test_manifest_and_import_sources(self):
        manifest = json.loads((ROOT/'SOURCE_MANIFEST.json').read_text(encoding='utf-8-sig'))
        self.assertEqual(manifest['release'], 'G1R19')
        sources = {p.name:p for p in (ROOT/'src/vba').glob('*.bas')}
        self.assertEqual(len(sources), 62)
        self.assertEqual(set(sources), set(manifest['modules']))
        for name,path in sources.items():
            record = manifest['modules'][name]
            imported = ROOT/'import'/name
            self.assertEqual(hashlib.sha256(path.read_bytes()).hexdigest(), record['utf8_sha256'], name)
            self.assertEqual(hashlib.sha256(imported.read_bytes()).hexdigest(), record['cp932_sha256'], name)
            utf8 = path.read_bytes().decode('utf-8').replace('\r\n','\n')
            cp932 = imported.read_bytes().decode('cp932').replace('\r\n','\n')
            self.assertEqual(utf8, cp932, name)
            self.assertTrue(utf8.startswith(f'Attribute VB_Name = "{path.stem}"'), name)

    def test_current_hydrology_and_cache_fixes_are_retained(self):
        for module,version in [('RPX_LowerCache','G1R18_Issue12'), ('RPX_SeepUI','G1R18_Issue15')]:
            actual = (ROOT/'src/vba'/f'{module}.bas').read_text(encoding='utf-8')
            expected = (VERSIONS/version/'src/vba'/f'{module}.bas').read_text(encoding='utf-8')
            self.assertEqual(actual, expected, module)
        fs = (ROOT/'src/vba/RPX_Fs.bas').read_text(encoding='utf-8')
        self.assertIn('XOriginalCompiled.subdivisions <> RSubdivisions', fs)
        self.assertIn('XOriginalCompiled.inputEpoch <> RPX_InputEpoch', fs)

    def test_accuracy_calls_use_the_unified_dll(self):
        source = (ROOT/'src/vba/RPX_AccuracyNative.bas').read_text(encoding='utf-8')
        self.assertNotIn('rpfem_accuracy_x64.dll', source)
        self.assertIn('rpfem_accuracy_abi_version', source)
        self.assertIn('Lib "rpfem_native_x64.dll"', source)
        self.assertEqual([p.name for p in (ROOT/'delivery').glob('*.dll')], ['rpfem_native_x64.dll'])
        manifest = json.loads((ROOT/'native/unified/validation/final_verification_manifest.json').read_text(encoding='utf-8-sig'))
        self.assertEqual(hashlib.sha256((ROOT/'delivery/rpfem_native_x64.dll').read_bytes()).hexdigest(), manifest['final_dll_sha256'])

    def test_packaged_distribution_has_one_runtime_dll(self):
        path = ROOT/'delivery/RPFEM_20261015_G1R19.zip'
        if not path.is_file():
            self.skipTest('Distribution ZIP is checked after packaging')
        with zipfile.ZipFile(path) as archive:
            self.assertIsNone(archive.testzip())
            dlls = [name for name in archive.namelist() if name.lower().endswith('.dll')]
            self.assertEqual(len(dlls), 1, dlls)
            self.assertTrue(dlls[0].endswith('/rpfem_native_x64.dll') or dlls[0] == 'rpfem_native_x64.dll')
            self.assertEqual(hashlib.sha256(archive.read(dlls[0])).hexdigest(), hashlib.sha256((ROOT/'delivery/rpfem_native_x64.dll').read_bytes()).hexdigest())


if __name__ == '__main__':
    unittest.main()
