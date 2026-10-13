"""Check rebuilt unchanged G1R17 DLL: ABI, exports and existing compiler oracle."""
import ctypes
import hashlib
import json
from pathlib import Path
import struct
import sys

ROOT = Path(__file__).resolve().parents[1]

def exports(path):
    data = path.read_bytes()
    pe = struct.unpack_from('<I', data, 0x3c)[0]
    assert data[pe:pe+4] == b'PE\0\0'
    machine, count = struct.unpack_from('<HH', data, pe+4)
    assert machine == 0x8664
    size = struct.unpack_from('<H', data, pe+20)[0]
    opt = pe+24
    assert struct.unpack_from('<H', data, opt)[0] == 0x20b
    er = struct.unpack_from('<I', data, opt+112)[0]
    sections = []
    for i in range(count):
        s = opt+size+40*i
        vs, va, rs, raw = struct.unpack_from('<IIII', data, s+8)
        sections.append((va, max(vs, rs), raw))
    def offset(rva):
        for va, length, raw in sections:
            if va <= rva < va+length:
                return raw+rva-va
        raise ValueError('RVA outside sections')
    entry = offset(er)
    n = struct.unpack_from('<I', data, entry+24)[0]
    names = offset(struct.unpack_from('<I', data, entry+32)[0])
    result = set()
    for i in range(n):
        at = offset(struct.unpack_from('<I', data, names+4*i)[0])
        result.add(data[at:data.index(b'\0', at)].decode('ascii'))
    return result

def main():
    rebuilt = Path(sys.argv[1]).resolve()
    original = ROOT.parent/'G1R17/delivery/rpfem_native_x64.dll'
    native_source = ROOT.parent/'G1R17/native'
    for p in native_source.iterdir():
        if p.suffix in ('.c', '.cpp', '.h', '.asm'):
            assert p.read_bytes() == (rebuilt.parent.parent/'native'/p.name).read_bytes(), p.name
    a, b = exports(original), exports(rebuilt)
    assert a == b and len(a) == 94
    dll = ctypes.WinDLL(str(rebuilt))
    assert dll.rpx_abi() == 0x00170005
    sys.path.insert(0, str(ROOT.parent/'G1R17/tools'))
    import test_compile_python
    test_compile_python.ROOT = rebuilt.parent.parent
    test_compile_python.run()
    result = {
        'abi': '0x00170005', 'exports_equal': True, 'exports_count': len(a),
        'source_unchanged': True, 'compiler_regression_cases': 120,
        'original_sha256': hashlib.sha256(original.read_bytes()).hexdigest(),
        'rebuilt_sha256': hashlib.sha256(rebuilt.read_bytes()).hexdigest()
    }
    (ROOT/'results/issue8_native.json').write_text(json.dumps(result, indent=2)+'\n', encoding='utf8')
    print(json.dumps(result))

if __name__ == '__main__':
    main()
