"""Read native RP25M001 snapshots without Excel or spreadsheet rewriting.

The read_state function preserves the binary layout and assertions of the
Native25 audit_batch reader. Only its Mesh import is made self-contained.
"""
import numpy as np
from model_support.schema import Mesh


def read_state(path, policy, phase):
    raw = path.read_bytes(); assert raw[:8] == b'RP25M001'; pos = 8
    def take(dtype, count):
        nonlocal pos
        out = np.frombuffer(raw, dtype=dtype, count=count, offset=pos).copy()
        pos += out.nbytes
        return out
    nn, ne, nm, nr, q = map(int, take('<i4', 5)); assert nr == 0
    scale, factor, fs, objective, native_audit = take('<f8', 5)
    xy = take('<f8', nn*2).reshape(nn, 2)
    tr = take('<i4', ne*5).reshape(ne, 5)
    mats = take('<f8', nm*4).reshape(nm, 4)
    body = take('<f8', ne*4).reshape(ne, 2, 2)
    nb = int(take('<i4', 1)[0]); bd = {}; kinds = ['load', 'fixed', 'roller_x', 'roller_y']
    for _ in range(nb):
        a, b, k, r = map(int, take('<i4', 4))
        loads = take('<f8', 4).reshape(2, 2).T
        bd[tuple(sorted((a, b)))] = {'kind': kinds[k], 'rigid': r, 'loads': loads}
    nf = int(take('<i4', 1)[0]); x = take('<f8', nf); assert pos == len(raw)
    assert phase == 'reference_total' and factor == fs and np.all(tr[:, 4] < 0)
    mesh = Mesh(xy, tr[:, :3], mats, tr[:, 3], tr[:, 4], body, bd,
                float(scale), 'native_batch25', load_phase=phase,
                provenance={'psi_policy': policy}).validate()
    return mesh, x, {'nodes': nn, 'elements': ne, 'q': q, 'factor': factor,
                     'fs': fs, 'objective': objective,
                     'native_audit': native_audit, 'scale': scale}
