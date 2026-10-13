"""Private original lower guards and positive dual coordinates.

All repository Python and every Python file in this new package are frozen
before repository imports. This module starts no solver or source audit by
itself. It does not import or modify any earlier upper/lower wrapper.
"""
from pathlib import Path
import hashlib
import importlib
import importlib.util
import json
import math
import shutil
import sys
import time

sys.dont_write_bytecode = True
REPO = Path(r'D:\Work\excel-vba-rpfem')
CODE = REPO / 'validation/accuracy_20261013/python'
CATALOG = REPO / 'docs/benchmarks/G1R17_native26_20261013/data/inputs'
GEOMETRY = REPO / 'versions/G1R17/results/geometry25'
HERE = Path(__file__).resolve().parent


def require(condition, message):
    if not condition:
        raise ValueError(message)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def hashes(paths):
    return {str(Path(path).resolve()): sha(path) for path in paths}


def write_json(path, value):
    Path(path).parent.mkdir(parents=True, exist_ok=True)
    Path(path).write_text(json.dumps(value, indent=2, allow_nan=False) + '\n', encoding='utf8')


def positive(value, label):
    require(math.isfinite(value) and value > 0, label + ' must be finite and positive')


def case_name(value):
    return {'N': 'P06N', 'A': 'P06A'}.get(value.upper(), value)


def validate_paths(args):
    require(args.case and '/' not in args.case and '\\' not in args.case
            and args.case not in ('.', '..'), 'Case must be a single catalog name')
    seed, output = args.seed.resolve(), args.output.resolve()
    require(seed != output and not output.is_relative_to(seed), 'Output must be outside the source seed')
    require(not output.exists(), 'Output must be a new directory')
    return seed, output


def scalar(data, key, expected):
    require(key in data.files and data[key].shape == () and data[key].item() == expected,
            'NPZ scalar identity differs: ' + key)


def loaded_repo_sources():
    paths = set()
    for module in list(sys.modules.values()):
        name = getattr(module, '__file__', None)
        if not name:
            continue
        path = Path(name).resolve()
        if path.suffix in ('.pyc', '.pyo'):
            try:
                path = Path(importlib.util.source_from_cache(str(path))).resolve()
            except ValueError:
                path = path.with_suffix('.py')
        if path.suffix == '.py' and path.is_relative_to(REPO):
            paths.add(path)
    return sorted(paths)


class Frozen:
    def __init__(self, args, guide=None):
        self.seed, self.output = validate_paths(args)
        self.case = args.case
        self.primal_paths = [self.seed / name for name in
                             ('best_lower_certificate.json', 'best_lower_mesh.npz', 'best_lower_field.npz')]
        self.input_paths = self.primal_paths + ([] if guide is None else [guide.resolve()])
        self.model_paths = [CATALOG / f'{args.case}.json', GEOMETRY / f'{args.case}_new_vba_mesh.bin']
        self.package_paths = sorted(HERE.glob('*.py'))
        self.code_paths = sorted(CODE.rglob('*.py')) + self.package_paths
        self.before = {'input': hashes(self.input_paths), 'model': hashes(self.model_paths),
                       'code': hashes(self.code_paths), 'package': hashes(self.package_paths)}
        self.output.mkdir(parents=True, exist_ok=False)
        write_json(self.output / 'sourcefreeze_before.json', self.before)
        self.copies = {}
        for path in self.primal_paths:
            self._copy(path, self.output / 'source_inputs/primal' / path.name)
        if guide is not None:
            self.guide_copy = self.output / 'source_inputs/guide' / guide.name
            self._copy(guide.resolve(), self.guide_copy)
        else:
            self.guide_copy = None
        self.catalog_copy = self.output / 'source_model' / self.model_paths[0].name
        for path in self.model_paths:
            self._copy(path, self.output / 'source_model' / path.name)
        for path in self.code_paths:
            relative = Path('repo_python') / path.relative_to(CODE) if path.is_relative_to(CODE) else Path('new_package') / path.name
            self._copy(path, self.output / 'source_code' / relative)
        self.check('after_snapshot_before_repository_import')
        sys.path.insert(0, str(CODE))
        self.core = importlib.import_module('experiment_core')
        self.driver = importlib.import_module('accuracy_driver')
        self.original_dual = importlib.import_module('original_dual')
        require(all(self.before['code'].get(str((REPO / path).resolve())) == digest
                    for path, digest in self.driver.SOURCE_HASHES.items()),
                'Import-time repository code differs from frozen bytes')
        self.check('after_repository_import')

    def _copy(self, path, target):
        path = path.resolve()
        expected = next(group[str(path)] for group in self.before.values() if str(path) in group)
        target.parent.mkdir(parents=True, exist_ok=True)
        shutil.copyfile(path, target)
        require(sha(path) == expected == sha(target), 'Live source changed during snapshot copy: ' + str(path))
        self.copies[str(path)] = str(target.resolve())

    def check(self, stage):
        require(sorted(HERE.glob('*.py')) == self.package_paths, 'New codepackage file inventory changed at ' + stage)
        after = {'input': hashes(self.input_paths), 'model': hashes(self.model_paths),
                 'code': hashes(self.code_paths), 'package': hashes(self.package_paths)}
        require(after == self.before, 'Live source/input/model/codepackage changed at ' + stage)
        require(all(sha(copy) == next(group[source] for group in self.before.values() if source in group)
                    for source, copy in self.copies.items()), 'Immutable snapshot changed at ' + stage)
        loaded = loaded_repo_sources()
        require(all(str(path) in self.before['code'] for path in loaded), 'Unfrozen repository Python imported at ' + stage)
        return {'stage': stage, 'sha256': after, 'sourcefreeze_exact': True,
                'snapshots': self.copies.copy(),
                'loaded_repo_python_sha256': {str(path): after['code'][str(path)] for path in loaded}}

    def load_source(self):
        core, np = self.core, self.core.np
        record = json.loads(self.catalog_copy.read_text(encoding='utf8'))
        require(record.get('case') == self.case and record['source_case'].get('id') == self.case,
                'Source model case differs')
        require(not record['source_case'].get('pore'), 'Pore model needs a coupled snapshot')
        policy = record['source_case']['policy']
        metadata = json.loads(Path(self.copies[str(self.primal_paths[0].resolve())]).read_text(encoding='utf8'))
        with np.load(self.copies[str(self.primal_paths[1].resolve())], allow_pickle=False) as data:
            xy, tri, ids = (data[key].copy() for key in ('xy', 'tri', 'material_id'))
        require(xy.ndim == 2 and xy.shape[1:] == (2,) and np.isfinite(xy).all()
                and tri.ndim == 2 and tri.shape[1:] == (3,) and tri.dtype.kind in 'iu'
                and ids.shape == (len(tri),) and ids.dtype.kind in 'iu', 'Saved source mesh layout differs')
        mesh = core.mesh_from_arrays(xy, tri, ids, record)
        require(mesh.provenance['case'] == self.case and mesh.provenance['psi_policy'] == policy,
                'Source mesh case or policy differs')
        with np.load(self.copies[str(self.primal_paths[2].resolve())], allow_pickle=False) as data:
            scalar(data, 'mode', 'lower')
            scalar(data, 'q', 8)
            require(data['fs'].shape == (), 'Source Fs must be scalar')
            fs = float(data['fs'].item())
            positive(fs, 'Source Fs')
            for key, expected in {'case': self.case, 'mesh_id': mesh.mesh_id,
                                  'physical_model_id': mesh.physical_model_id, 'policy': policy}.items():
                if key in data.files:
                    scalar(data, key, expected)
            # Never read stored primal optimization duals or auxiliary arrays.
            field = data['x'].copy()
        identity = {'case': self.case, 'mode': 'lower', 'q': 8, 'fs': fs,
                    'mesh_id': mesh.mesh_id, 'physical_model_id': mesh.physical_model_id, 'policy': policy}
        require(all(metadata.get(key) == value for key, value in identity.items() if key != 'policy'),
                'Exact source certificate/field/mesh/Fs/model identity differs')
        if 'policy' in metadata:
            require(metadata['policy'] == policy, 'Source certificate policy differs')
        prepared, digest, layout = original_lower(core, self.original_dual, mesh, fs)
        require(field.shape == (18 * len(mesh.tri) + 1,) == prepared.original[-1].shape
                and np.isfinite(field).all(), 'Finite full original free-load source field required')
        audit = core.evaluate(prepared, field, None)
        require(audit['physical_pass'] and audit['witness'], 'Fresh full original q8 lower source witness gate failed')
        require(digest == program_hash(core, prepared.original), 'Source original program changed during audit')
        audit.update(identity, source_flags_ignored=True, origin='fresh_original_lower_q8_source_witness')
        self.check('after_fresh_original_source_witness_audit')
        return mesh, field, fs, identity, audit, prepared, digest, layout

    def provenance_packets(self):
        return {
            'source_witness_sha256_json': {path.name: self.before['input'][str(path.resolve())] for path in self.primal_paths},
            'source_model_sha256_json': self.before['model'],
            'source_code_sha256_json': self.before['code'],
            'new_package_sha256_json': self.before['package'],
        }


def program_hash(core, program):
    digest = hashlib.sha256()
    for value in program:
        arrays = (core.np.asarray(value.shape, dtype='<i8'), value.indptr,
                  value.indices, value.data) if core.sparse.issparse(value) else (core.np.asarray(value),)
        for array in arrays:
            digest.update(str(array.dtype).encode())
            digest.update(core.np.asarray(array.shape, dtype='<i8').tobytes())
            digest.update(core.np.ascontiguousarray(array).tobytes())
    return digest.hexdigest()


def original_lower(core, original_dual, mesh, fs):
    prepared = core.Prepared(mesh, fs, 'lower', 8, reduction='none', fixed_load=None)
    E, b, G, h, dims, cost = prepared.original
    expected = core.np.zeros(18 * len(mesh.tri) + 1)
    expected[-1] = -1.
    require(prepared.mode == 'lower' and prepared.q == 8 and prepared.kind == 'none' and prepared.T is None
            and core.np.array_equal(cost, expected)
            and core.np.array_equal(cost, prepared.physical_original_objective)
            and core.np.array_equal(b, core.np.zeros_like(b))
            and (E.shape, G.shape, len(b), len(h), tuple(dims)) == prepared.physical_original_layout
            and all(d == 3 for d in dims) and original_dual.eligibility(prepared) is None,
            'Original free-load lower q8 NONE program, cost[-1]=-1, b=0 and full physical layout/SOCs are required')
    layout = {'E_shape': E.shape, 'G_shape': G.shape, 'b_shape': b.shape, 'h_shape': h.shape,
        'cost_shape': cost.shape, 'full_original_dual_shape': [len(b) + len(h)],
        'cone_count': len(dims), 'cone_dimensions': sorted(set(dims)),
        'free_load_last_cost': float(cost[-1]), 'b_exactly_zero': True,
        'h_nonzero_count': int(core.np.count_nonzero(h)), 'fixed_load': None, 'reduction': 'none'}
    for matrix in (E, G):
        for array in (matrix.data, matrix.indices, matrix.indptr):
            array.setflags(write=False)
    for array in (b, h, cost):
        array.setflags(write=False)
    prepared.original = E, b, G, h, tuple(dims), cost
    return prepared, program_hash(core, prepared.original), layout


def coordinate_scales(core, prepared, minimum, maximum):
    np = core.np
    E, _, G, _, dims, _ = prepared.original
    y_magnitude = np.asarray(abs(E).max(axis=1).toarray()).ravel()
    z_magnitude = np.asarray(abs(G).max(axis=1).toarray()).ravel()
    raw_ky = 1. / np.maximum(1e-12, y_magnitude)
    ky = np.clip(raw_ky, minimum, maximum)
    kz = np.empty(G.shape[0], dtype=float)
    offset = 0
    for dimension in dims:
        magnitude = float(np.max(z_magnitude[offset:offset + dimension], initial=0.))
        kz[offset:offset + dimension] = 1. / max(1., magnitude)
        offset += dimension
    require(offset == len(kz) and ky.shape == (len(prepared.original[1]),)
            and np.isfinite(ky).all() and np.isfinite(kz).all()
            and np.all(ky > 0) and np.all(kz > 0), 'Coordinate scales must be finite positive in original dual layout')
    offset = 0
    for dimension in dims:
        block = kz[offset:offset + dimension]
        require(np.all(block == block[0]), 'One common positive Z scale is required per original SOC')
        offset += dimension
    report = {'Y_formula_before_cap': '1/max(1e-12,maxabs corresponding E.T column)',
        'Y_finite_cap': [minimum, maximum], 'Y_cap_low_count': int(np.count_nonzero(raw_ky < minimum)),
        'Y_cap_high_count': int(np.count_nonzero(raw_ky > maximum)),
        'Z_formula': '1/max(1,maxabs all G.T columns in each original product-cone block)',
        'Y_scale_min': float(np.min(ky)), 'Y_scale_max': float(np.max(ky)),
        'Z_scale_min': float(np.min(kz)), 'Z_scale_max': float(np.max(kz)),
        'finite_positive': True, 'Z_common_positive_scale_per_cone': True,
        'coordinate_change_is_bijective': True, 'scale_is_proof': False}
    ky.setflags(write=False)
    kz.setflags(write=False)
    return ky, kz, report


def transformed_constraints(core, prepared, ky, kz, threshold):
    """Build only auxiliary constraints; no cone objects or solver are created."""
    np, sparse = core.np, core.sparse
    E, b, G, h, dims, cost = prepared.original
    ny, nz = len(b), len(h)
    stationary = sparse.hstack([E.T @ sparse.diags(ky), -(G.T @ sparse.diags(kz))], format='csc')
    product_cone = sparse.hstack([sparse.csc_matrix((nz, ny)), -sparse.eye(nz, format='csc')], format='csc')
    # Both terms are retained. Original lower h is generally nonzero.
    separation = sparse.csc_matrix(np.r_[b * ky, h * kz].reshape(1, -1))
    matrix = sparse.vstack([stationary, product_cone, separation], format='csc')
    rhs = np.r_[-cost, np.zeros(nz), -threshold]
    require(np.isfinite(matrix.data).all() and np.isfinite(rhs).all(), 'Nonfinite transformed helper coefficients')
    return matrix, rhs


def solve_scaled_helper(core, original_dual, prepared, args):
    np, sparse, clarabel = core.np, core.sparse, core.clarabel
    E, b, G, h, dims, cost = prepared.original
    ky, kz, scale_report = coordinate_scales(core, prepared, args.y_scale_min, args.y_scale_max)
    ny, nz = len(b), len(h)
    matrix, rhs = transformed_constraints(core, prepared, ky, kz, original_dual.threshold(prepared))
    cones = ([clarabel.ZeroConeT(len(cost))]
        + [clarabel.NonnegativeConeT(d) if d == 1 else clarabel.SecondOrderConeT(d) for d in dims]
        + [clarabel.NonnegativeConeT(1)])
    P = (sparse.eye(ny + nz, format='csc') * args.regularization if args.regularization
         else sparse.csc_matrix((ny + nz, ny + nz)))
    options = {'equilibrate_enable': bool(args.equilibrate), 'time_limit': float(args.seconds),
        'tol_gap_abs': 1e-13, 'tol_gap_rel': 1e-13, 'tol_feas': 1e-13,
        'tol_infeas_abs': 1e-13, 'tol_infeas_rel': 1e-13,
        'iterative_refinement_max_iter': 30,
        'iterative_refinement_abstol': 1e-15, 'iterative_refinement_reltol': 1e-15}
    solver = clarabel.DefaultSolver(P, np.zeros(ny + nz), matrix, rhs, cones, core.settings(options, 350))
    started = time.perf_counter()
    solver.solve()
    solution = solver.get_solution()
    # Never read solution.z: only auxiliary primal coordinates solution.x
    # restore the full original optimization dual (Y,Z).
    new_coordinates = np.asarray(solution.x, dtype=float).copy()
    require(new_coordinates.shape == (ny + nz,) and np.isfinite(new_coordinates).all(), 'Finite exact-layout helper solution.x required')
    restored = np.r_[ky * new_coordinates[:ny], kz * new_coordinates[ny:]]
    require(np.isfinite(restored).all(), 'Nonfinite restored original dual')
    return restored, ky, kz, scale_report, {
        'status': str(solution.status), 'iterations': solution.iterations,
        'seconds': time.perf_counter() - started, 'settings': options, 'max_iter': 350,
        'auxiliary_regularization_new_coordinates': args.regularization,
        'helper_status_is_proof': False, 'auxiliary_solution_z_read_or_used': False,
        'transformed_constraint_dual_is_original_guide': False,
        'original_h_term_preserved': True,
        'stationary_equation': '[E.T*diag(ky),-G.T*diag(kz)]@(Ynew,Znew)=-physical_cost',
        'separation_equation': '[b*ky,h*kz]@(Ynew,Znew)<=-original_threshold',
    }


def reject(output, reason):
    """Quarantine this run's evidence while keeping source snapshots unchanged."""
    output = Path(output)

    def demote(value):
        if isinstance(value, dict):
            if 'root_complete' in value:
                value['root_complete'] = False
            for item in value.values():
                demote(item)
        elif isinstance(value, list):
            for item in value:
                demote(item)

    for directory in list(output.iterdir()):
        if not directory.is_dir() or directory.name in ('source_inputs', 'source_model', 'source_code'):
            continue
        for path in list(directory.rglob('*.json')):
            data = json.loads(path.read_text(encoding='utf8'))
            demote(data)
            if isinstance(data, dict):
                data.update(publication_rejected=True, publication_rejection_reason=reason)
            write_json(path, data)
            if path.name in ('best_lower_certificate.json', 'summary.json', 'import_audit.json'):
                path.rename(path.with_name('rejected_' + path.name))
    for name in ('original_helper_full_dual.npz', 'helper_audit.json', 'closure_audit.json'):
        path = output / name
        if path.exists():
            if path.suffix == '.json':
                data = json.loads(path.read_text(encoding='utf8'))
                demote(data)
                data.update(publication_rejected=True, publication_rejection_reason=reason)
                write_json(path, data)
            path.rename(path.with_name('rejected_' + name))
    write_json(output / 'publication_rejected.json', {'accepted': False, 'root_complete': False,
        'publishes_primal_bound': False, 'reason': reason, 'source_flags_ignored': True})
