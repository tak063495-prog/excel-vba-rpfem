"""Compare a VBA fixed-point original program with Python; never solve.

Geometry is imported, not archived stresses/velocities/roots. A native primal,
when present, is audited in a newly assembled canonical Python q8 program.
No full root or complete adaptive-run result follows from this checkpoint.
"""
from pathlib import Path
import argparse, hashlib, json, os, sys
for name in ('OMP_NUM_THREADS', 'OPENBLAS_NUM_THREADS', 'MKL_NUM_THREADS'):
    os.environ[name] = '1'
REPO = Path(__file__).resolve().parents[3]
sys.path.insert(0, str(REPO/'validation/G1R19_python26_20261016/python'))
import benchmark_models as models
import experiment_core as core
np, sparse = core.np, core.sparse


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


class Reader:
    def __init__(self, path, magic):
        self.raw = Path(path).read_bytes(); self.position = len(magic)
        if self.raw[:len(magic)] != magic:
            raise ValueError('Unexpected binary format')
    def take(self, dtype, count):
        value = np.frombuffer(self.raw, dtype=dtype, count=count, offset=self.position).copy()
        self.position += value.nbytes
        return value
    def end(self):
        if self.position != len(self.raw):
            raise ValueError('Unexpected trailing data')


def geometry(path):
    r = Reader(path, b'RP25G001')
    nn, ne, nm, nr, q = map(int, r.take('<i4', 5)); scale = float(r.take('<f8', 1)[0])
    if nr != 0 or q != 8 or min(nn, ne, nm) <= 0:
        raise ValueError('Dry soil original-q8 geometry required')
    xy = r.take('<f8', 2*nn).reshape(nn, 2)
    cells = r.take('<i4', 5*ne).reshape(ne, 5)
    materials = r.take('<f8', 4*nm).reshape(nm, 4)
    body = r.take('<f8', 4*ne).reshape(ne, 2, 2)
    boundaries = {}
    for _ in range(int(r.take('<i4', 1)[0])):
        a, b, kind, rigid = map(int, r.take('<i4', 4))
        edge = tuple(sorted((a, b)))
        if edge in boundaries or not 0 <= kind < 4 or rigid != -1:
            raise ValueError('Invalid exterior boundary')
        boundaries[edge] = dict(kind=['load','fixed','roller_x','roller_y'][kind], rigid=rigid,
                               loads=r.take('<f8', 4).reshape(2, 2).T)
    r.end()
    return xy, cells, materials, body, boundaries, scale


def original_program(path):
    r = Reader(path, b'P06VBA01')
    n, neq, nc = map(int, r.take('<i4', 3)); cost = r.take('<f8', n)
    ptr, ids, values, rhs, dims = [0], [], [], [], []
    def row():
        count = int(r.take('<i4', 1)[0]); rhs.append(float(r.take('<f8', 1)[0]))
        data = r.take(np.dtype([('id','<i4'), ('value','<f8')]), count)
        ids.append(data['id']); values.append(data['value']); ptr.append(ptr[-1]+count)
    for _ in range(neq):
        row()
    for _ in range(nc):
        dimension = int(r.take('<i4', 1)[0]); dims.append(dimension)
        if dimension not in (1, 3):
            raise ValueError('Unsupported original cone')
        for _ in range(dimension):
            row()
    r.end()
    a = sparse.csr_matrix((np.concatenate(values), np.concatenate(ids), ptr), shape=(len(ptr)-1, n))
    rhs = np.asarray(rhs); e, g = a[:neq], a[neq:]
    norms = np.sqrt(np.asarray(e.power(2).sum(axis=1)).ravel())
    if not np.all(norms > 0):
        raise ValueError('Invalid equality row')
    return sparse.diags(1/norms)@e, rhs[:neq]/norms, g, rhs[neq:], dims, cost


def difference(left, right):
    if left.shape != right.shape:
        return dict(equal_shape=False, native_shape=list(left.shape), python_shape=list(right.shape), pass_=False)
    error = float(abs(left-right).max()) if sparse.issparse(left) else float(np.max(abs(left-right), initial=0.))
    return dict(equal_shape=True, max_absolute_difference=error, tolerance=1e-10, pass_=error <= 1e-10)


def audit(folder, *, assembly_only=False):
    record = json.loads((folder/'summary.json').read_text(encoding='utf8'))
    case, side, fs = record['case'], record['side'], record['comparison_fs']
    geo = REPO/'validation/G1R20_native_geometry_replay_20261016/geometry'/(case+'_'+side+'.bin')
    models.install_model_builder(); canonical = models._record(case)
    xy, cells, materials, body, boundaries, scale = geometry(geo)
    mesh = core.mesh_from_arrays(xy, cells[:,:3], cells[:,3], canonical)
    assert np.array_equal(materials, mesh.materials) and np.array_equal(body, mesh.body)
    assert mesh.stress_scale == scale and boundaries.keys() == mesh.boundary.keys()
    for edge, boundary in boundaries.items():
        for key in ('kind', 'rigid'):
            assert boundary[key] == mesh.boundary[edge][key]
        assert np.array_equal(boundary['loads'], mesh.boundary[edge]['loads'])
    prepared = core.Prepared(mesh, fs, side, 8, reduction='none')
    exported = original_program(folder/'fixed_original_assembled.bin')
    comparisons = {key: difference(left, right) for key, left, right in
                   zip(('E','b','G','h','cost'), exported[:4]+(exported[5],), prepared.original[:4]+(prepared.original[5],))}
    comparisons['cone_dimensions'] = dict(pass_=np.array_equal(exported[4], prepared.original[4]))
    result = dict(case=case, side=side, fs=fs, q=8, elements=len(mesh.tri), nodes=len(mesh.xy),
                  scope='Same geometry/fixed Fs original-program and optional primal checkpoint only; no own root or full adaptive-run claim',
                  optimizer_calls=0, comparisons=comparisons, canonical_material_body_boundary_exact=True,
                  physical_model_id=mesh.physical_model_id, python_policy=mesh.provenance['psi_policy'],
                  source_sha256={str(p.relative_to(REPO)):sha(p) for p in (geo, folder/'fixed_original_assembled.bin', Path(__file__))})
    passed = all(v['pass_'] for v in comparisons.values())
    if not assembly_only:
        native = models._read_native_state(folder/'fixed_native_primal.bin')
        assert native['fs'] == fs and native['q'] == 8
        assert np.array_equal(native['xy'], xy) and np.array_equal(native['cells'], cells)
        assert np.array_equal(native['materials'], materials) and np.array_equal(native['body'], body)
        before = hashlib.sha256(native['field'].tobytes()).hexdigest()
        evaluated = core.evaluate(prepared, native['field'])
        assert before == hashlib.sha256(native['field'].tobytes()).hexdigest()
        result.update(independent_original_q8=evaluated, primal_sha256=sha(folder/'fixed_native_primal.bin'),
                      primal_unchanged=True, archived_primal_imported=False, root_complete=False)
        passed = passed and bool(evaluated['physical_pass'] and evaluated['witness'])
    result['status'] = ('ASSEMBLY_PARITY_PASS' if assembly_only else 'FIXED_POINT_PARITY_PASS') if passed else 'FAIL'
    return result


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--run', type=Path, required=True); parser.add_argument('--output', type=Path, required=True)
    parser.add_argument('--assembly-only', action='store_true'); args = parser.parse_args()
    if args.output.exists():
        raise FileExistsError('Preserve existing audit')
    result = audit(args.run.resolve(), assembly_only=args.assembly_only)
    args.output.write_text(json.dumps(result, indent=2, ensure_ascii=False, allow_nan=False)+'\n', encoding='utf8')
    print(json.dumps(result, ensure_ascii=False))
    raise SystemExit(0 if result['status'].endswith('_PASS') else 2)
