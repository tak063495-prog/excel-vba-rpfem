"""Standard-library integrity check for the separate Issue #8 patch artifacts."""
import hashlib
import json
from pathlib import Path
import zipfile

ROOT = Path(__file__).resolve().parents[1]

def main():
    manifest = json.loads((ROOT/'MANIFEST.json').read_text(encoding='utf8'))
    for name, record in manifest['files'].items():
        data = (ROOT/name).read_bytes()
        assert len(data) == record['bytes'], name
        assert hashlib.sha256(data).hexdigest() == record['sha256'], name
    for name in ['RPX_Seep.bas', 'RPX_SeepUI.bas']:
        assert (ROOT/'import'/name).read_bytes().decode('cp932').replace('\r\n','\n') == (ROOT/'src/vba'/name).read_text(encoding='utf8')
    original = ROOT.parent/'G1R17/delivery/rpfem_native_x64.dll'
    assert original.read_bytes() == (ROOT/'delivery/rpfem_native_x64.dll').read_bytes()
    for p in (ROOT/'delivery').glob('*.xlsm'):
        with zipfile.ZipFile(p) as z:
            assert z.testzip() is None
            assert 'xl/vbaProject.bin' in z.namelist()
    print('ISSUE8 PACKAGE INTEGRITY PASS', len(manifest['files']))

if __name__ == '__main__':
    main()
