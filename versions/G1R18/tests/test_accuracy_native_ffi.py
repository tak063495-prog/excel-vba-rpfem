"""Independent x64 C ABI and original q8 audits for the G1R18 owned worker.

Run with unittest discovery. No solver status creates a bound in these tests:
native full fields are audited against the original Python E/G and physics.
The DLL must already be built; RPFEM_ACCURACY_DLL can select a build artifact.
"""
from __future__ import annotations

import copy
import ctypes as C
import hashlib
import json
import math
import os
from pathlib import Path
import sys
import time
import unittest

REPO = Path(__file__).resolve().parents[3]
NATIVE = REPO / 'versions/G1R18/native/accuracy'
PYTHON = REPO / 'validation/accuracy_20261013/python'
sys.path.insert(0, str(PYTHON))
import experiment_core as core
import accuracy_driver as driver
import c0_slope_wedge_driver as wedge
from original_dual import audit_original_dual

np, sparse, clarabel = core.np, core.sparse, core.clarabel
I32, U32, U64, F64 = C.c_int32, C.c_uint32, C.c_uint64, C.c_double
PI, PD = C.POINTER(I32), C.POINTER(F64)
OK, INVALID, LIMIT, HANDLE, STATE, BUFFER, BUSY = 0, -1, -2, -3, -4, -5, -8
CREATED, RUNNING, DONE, CANCELLED, ERROR = range(5)


class Problem(C.Structure):
    _fields_ = [(name, U32) for name in
                ('struct_size', 'abi_version', 'n', 'm', 'nnz', 'cone_count')] + [
        ('row_ptr', PI), ('col_index', PI), ('A_values', PD), ('b', PD),
        ('Pdiag', PD), ('q', PD), ('cone_kind', PI), ('cone_dim', PI)]


class Settings(C.Structure):
    _fields_ = [(name, U32) for name in
                ('struct_size', 'abi_version', 'max_iter', 'max_threads')] + [
        (name, F64) for name in ('time_limit', 'tol_gap_abs', 'tol_gap_rel',
            'tol_feas', 'tol_infeas_abs', 'tol_infeas_rel',
            'iterative_refinement_reltol', 'iterative_refinement_abstol')] + [
        ('iterative_refinement_max_iter', U32), ('equilibrate', U32),
        ('static_regularization_constant', F64), ('dynamic_regularization_eps', F64),
        ('dynamic_regularization_delta', F64), ('static_regularization_enable', U32),
        ('dynamic_regularization_enable', U32), ('direct_method', U32), ('reserved', U32)]


class Progress(C.Structure):
    _fields_ = [(name, U32) for name in ('struct_size', 'abi_version', 'state',
        'solver_status', 'iterations', 'cancel_requested', 'n', 'm')] + [
        (name, F64) for name in ('solve_time', 'primal_objective', 'dual_objective',
            'primal_residual', 'dual_residual', 'diagnostics_valid')]


class EqualityProblem(C.Structure):
    _fields_ = [(name, U32) for name in
                ('struct_size', 'abi_version', 'n', 'm', 'nnz', 'reserved')] + [
        ('row_ptr', PI), ('col_index', PI), ('A_values', PD), ('b', PD), ('x0', PD)]


class EqualityReport(C.Structure):
    _fields_ = [(name, U32) for name in
                ('struct_size', 'abi_version', 'istop', 'iterations')] + [
        (name, F64) for name in ('normr', 'normar', 'norma', 'conda', 'normx')]


def progress():
    return Progress(struct_size=C.sizeof(Progress), abi_version=1)


def sha(path):
    return hashlib.sha256(Path(path).read_bytes()).hexdigest()


class OwnedInput:
    """Keep actual, correctly allocated C arrays alive during create only."""
    def __init__(self, A, b, q, kinds, dims, p=None):
        A = sparse.csr_matrix(A, dtype=float).copy()
        A.sum_duplicates(); A.sort_indices()
        self.arrays = {
            'row_ptr': np.asarray(A.indptr, np.int32).copy(),
            'col_index': np.asarray(A.indices, np.int32).copy(),
            'A_values': np.asarray(A.data, np.float64).copy(),
            'b': np.asarray(b, np.float64).copy(),
            'Pdiag': np.zeros(A.shape[1]) if p is None else np.asarray(p, np.float64).copy(),
            'q': np.asarray(q, np.float64).copy(),
            'cone_kind': np.asarray(kinds, np.int32).copy(),
            'cone_dim': np.asarray(dims, np.int32).copy()}
        self.problem = Problem(struct_size=C.sizeof(Problem), abi_version=1,
            n=A.shape[1], m=A.shape[0], nnz=A.nnz, cone_count=len(dims))
        for name, a in self.arrays.items():
            setattr(self.problem, name, a.ctypes.data_as(PI if a.dtype == np.int32 else PD))

    @classmethod
    def original(cls, prepared, p=None, q=None):
        E, b, G, h, dims, cost = prepared.original
        return cls(sparse.vstack([E, -G]), np.r_[b, h], cost if q is None else q,
            [0] + [1 if d == 1 else 3 for d in dims], [len(b), *dims], p)


def physical_triangle():
    # c=1, phi=psi=0, gamma=1: original lower h is nonzero.
    boundary = {(0, 1): {'kind': 'fixed', 'rigid': -1, 'loads': np.zeros((2, 2))},
        (1, 2): {'kind': 'load', 'rigid': -1, 'loads': np.zeros((2, 2))},
        (0, 2): {'kind': 'load', 'rigid': -1, 'loads': np.zeros((2, 2))}}
    return core.Mesh(np.array([[0., 0.], [1., 0.], [0., 1.]]),
        np.array([[0, 1, 2]], np.int32), np.array([[1., 0., 0., 1.]]),
        np.array([0], np.int32), np.array([-1], np.int32),
        np.array([[[0., 0.], [0., -1.]]]), boundary, 1.,
        'ffi_real_cohesive_triangle', provenance={'psi_policy': 'hold_initial',
        'case': 'TEST_FFI_COHESIVE'}).validate()


class NativeAccuracyFFITests(unittest.TestCase):
    records = []

    @classmethod
    def setUpClass(cls):
        if os.name != 'nt' or C.sizeof(C.c_void_p) != 8:
            raise unittest.SkipTest('Windows x64 native worker test')
        cls.dll_path = Path(os.environ.get('RPFEM_ACCURACY_DLL',
            str(NATIVE / 'bin/rpfem_accuracy_x64.dll'))).resolve()
        if not cls.dll_path.is_file():
            raise RuntimeError(f'Build the G1R18 accuracy DLL first: {cls.dll_path}')
        cls.protected = [cls.dll_path, NATIVE / 'src/lib.rs',
            NATIVE / 'src/equality.rs', NATIVE / 'LICENSE_LSMR.txt',
            NATIVE / 'rpfem_accuracy.h', NATIVE / 'Cargo.toml', NATIVE / 'Cargo.lock',
            Path(__file__), PYTHON / 'experiment_core.py', PYTHON / 'original_dual.py',
            Path(sys.modules['independent_audit'].__file__), PYTHON / 'c0_slope_wedge_driver.py',
            PYTHON / 'vendor/model_support/lower_assembly.py',
            PYTHON / 'vendor/model_support/upper_assembly.py', driver.CATALOG / 'D02A.json']
        cls.before = {str(p): sha(p) for p in cls.protected}
        cls.dll = C.CDLL(str(cls.dll_path))
        signatures = {
            'abi_version': (U32, []), 'default_settings': (I32, [C.POINTER(Settings)]),
            'create': (I32, [C.POINTER(Problem), C.POINTER(Settings), C.POINTER(U64)]),
            'start': (I32, [U64]), 'poll': (I32, [U64, C.POINTER(Progress)]),
            'cancel': (I32, [U64]),
            'result': (I32, [U64, PD, U32, PD, U32, C.POINTER(Progress)]),
            'equalities': (I32, [C.POINTER(EqualityProblem), U32, F64, F64, F64,
                PD, U32, C.POINTER(EqualityReport)]),
            'destroy': (I32, [U64])}
        for name, (restype, args) in signatures.items():
            fn = getattr(cls.dll, 'rpfem_accuracy_' + name)
            fn.restype, fn.argtypes = restype, args
            setattr(cls, name, staticmethod(fn))

    @classmethod
    def tearDownClass(cls):
        after = {str(p): sha(p) for p in cls.protected}
        report = {'dll': str(cls.dll_path), 'protected_sha256': cls.before,
            'protected_sha256_after': after, 'sources_unchanged': cls.before == after,
            'scope': 'small original q8 and ABI; no solver-status bound adoption',
            'fixtures': cls.records}
        destination = os.environ.get('RPFEM_ACCURACY_QA_REPORT')
        if destination:
            core.write(destination, report)
        if cls.before != after:
            raise AssertionError('Native artifact/source changed during independent QA')

    def config(self):
        settings = Settings()
        self.assertEqual(self.default_settings(C.byref(settings)), OK)
        settings.time_limit = 15.
        return settings

    def create_input(self, data, settings=None, expected=OK):
        handle = U64(0)
        rc = self.create(C.byref(data.problem), C.byref(settings or self.config()), C.byref(handle))
        self.assertEqual(rc, expected)
        if expected != OK:
            self.assertEqual(handle.value, 0, 'rejected input must not publish a handle')
        return handle.value

    def cleanup(self, handle):
        deadline = time.monotonic() + 25
        while True:
            rc = self.destroy(handle)
            if rc == OK:
                return
            self.assertEqual(rc, BUSY)
            self.assertLess(time.monotonic(), deadline, 'destroy must eventually release owned worker')
            time.sleep(.002)

    def await_terminal(self, handle):
        deadline = time.monotonic() + 25
        while True:
            p = progress()
            self.assertEqual(self.poll(handle, C.byref(p)), OK)
            if p.state != RUNNING:
                return p
            self.assertLess(time.monotonic(), deadline)
            time.sleep(.002)

    def run_input(self, data):
        handle = self.create_input(data)
        try:
            self.assertEqual(self.start(handle), OK)
            terminal = self.await_terminal(handle)
            self.assertEqual(terminal.state, DONE)
            x, z = np.zeros(data.problem.n), np.zeros(data.problem.m)
            p = progress()
            self.assertEqual(self.result(handle, x.ctypes.data_as(PD), len(x),
                z.ctypes.data_as(PD), len(z), C.byref(p)), OK)
            self.assertTrue(np.isfinite(x).all() and np.isfinite(z).all())
            self.assertEqual((p.n, p.m), (len(x), len(z)))
            return x, z, p
        finally:
            self.cleanup(handle)

    @staticmethod
    def tiny(p=None, q=None):
        return OwnedInput([[-1.]], [-1.], [1.] if q is None else q, [1], [1], p)

    def python_solution(self, data):
        a = data.arrays; p = data.problem
        A = sparse.csr_matrix((a['A_values'], a['col_index'], a['row_ptr']), shape=(p.m, p.n))
        cones = [clarabel.ZeroConeT(int(d)) if k == 0 else
            clarabel.NonnegativeConeT(int(d)) if k == 1 else clarabel.SecondOrderConeT(int(d))
            for k, d in zip(a['cone_kind'], a['cone_dim'])]
        s = core.settings({'time_limit': 15., 'equilibrate_enable': False,
            'tol_gap_abs': 1e-13, 'tol_gap_rel': 1e-13, 'tol_feas': 1e-13,
            'tol_infeas_abs': 1e-13, 'tol_infeas_rel': 1e-13,
            'iterative_refinement_reltol': 1e-15, 'iterative_refinement_abstol': 1e-15,
            'iterative_refinement_max_iter': 30}, 350)
        solver = clarabel.DefaultSolver(sparse.diags(a['Pdiag'], format='csc'), a['q'],
            A.tocsc(), a['b'], cones, s)
        solver.solve()
        solution = solver.get_solution()
        return np.asarray(solution.x), np.asarray(solution.z)

    def test_abi_layout_and_default_settings(self):
        self.assertEqual(self.abi_version(), 1)
        self.assertEqual([C.sizeof(t) for t in (Problem, Settings, Progress)], [88, 128, 80])
        self.assertEqual([Problem.row_ptr.offset, Problem.q.offset, Settings.time_limit.offset,
            Settings.static_regularization_constant.offset, Progress.solve_time.offset], [24, 64, 16, 88, 32])
        s = self.config()
        self.assertEqual((s.struct_size, s.abi_version, s.max_iter, s.max_threads,
            s.equilibrate), (128, 1, 350, 1, 0))
        self.assertEqual(s.tol_feas, 1e-13)
        self.assertEqual(self.default_settings(None), INVALID)

    def test_linear_and_diagonal_qp_match_python(self):
        for data, expected in [(self.tiny(), 1.), (self.tiny([1.], [-2.]), 2.)]:
            with self.subTest(expected=expected):
                x, z, _ = self.run_input(data)
                xp, zp = self.python_solution(data)
                np.testing.assert_allclose(x, xp, atol=1e-9, rtol=1e-9)
                np.testing.assert_allclose(z, zp, atol=1e-9, rtol=1e-9)
                self.assertAlmostEqual(x[0], expected, places=8)

    def test_create_copies_every_input_before_worker_starts(self):
        data = self.tiny()
        h = self.create_input(data)
        try:
            for a in data.arrays.values():
                a.fill(777)
            self.assertEqual(self.start(h), OK)
            self.assertEqual(self.await_terminal(h).state, DONE)
            x, z, p = np.zeros(1), np.zeros(1), progress()
            self.assertEqual(self.result(h, x.ctypes.data_as(PD), 1, z.ctypes.data_as(PD), 1, C.byref(p)), OK)
            self.assertAlmostEqual(x[0], 1., places=8)
            self.assertAlmostEqual(z[0], 1., places=8)
        finally:
            self.cleanup(h)

    def test_rejects_abi_null_alignment_and_dimension_limits(self):
        for name, value, expected in [('struct_size', 87, INVALID), ('abi_version', 2, INVALID),
            ('n', 0, LIMIT), ('n', 2_000_001, LIMIT), ('m', 8_000_001, LIMIT),
            ('nnz', 50_000_001, LIMIT), ('cone_count', 2, LIMIT)]:
            with self.subTest(name=name):
                data = self.tiny(); setattr(data.problem, name, value)
                self.create_input(data, expected=expected)
        data = self.tiny(); data.problem.q = PD()
        self.create_input(data, expected=INVALID)
        # Misalignment points inside a real allocation; no arbitrary invalid address is dereferenced.
        storage = C.create_string_buffer(32)
        data = self.tiny(); data.problem.q = C.cast(C.addressof(storage) + 1, PD)
        self.create_input(data, expected=INVALID)
        h = U64()
        self.assertEqual(self.create(None, C.byref(self.config()), C.byref(h)), INVALID)
        data = self.tiny()
        self.assertEqual(self.create(C.byref(data.problem), None, C.byref(h)), INVALID)

    def test_rejects_invalid_settings(self):
        for name, value in [('struct_size', 0), ('abi_version', 0), ('max_iter', 0),
            ('max_threads', 9), ('time_limit', math.inf), ('tol_feas', 0.),
            ('tol_gap_abs', math.nan), ('equilibrate', 2), ('direct_method', 3),
            ('static_regularization_constant', -1.), ('dynamic_regularization_eps', 0.),
            ('iterative_refinement_max_iter', 0), ('reserved', 1)]:
            with self.subTest(name=name):
                s = self.config(); setattr(s, name, value)
                self.create_input(self.tiny(), s, INVALID)

    def test_rejects_nonfinite_and_negative_pdiag(self):
        for name in ('A_values', 'b', 'Pdiag', 'q'):
            for bad in (math.nan, math.inf, -math.inf):
                with self.subTest(name=name, bad=bad):
                    data = self.tiny(); data.arrays[name][0] = bad
                    self.create_input(data, expected=INVALID)
        data = self.tiny(); data.arrays['Pdiag'][0] = -1.
        self.create_input(data, expected=INVALID)

    def test_rejects_csr_and_cone_structure(self):
        for name, value in [('row_ptr', [1, 1]), ('row_ptr', [0, 0]),
            ('row_ptr', [0, -1]), ('col_index', [-1]), ('col_index', [1]),
            ('cone_kind', [2]), ('cone_dim', [0]), ('cone_dim', [2])]:
            with self.subTest(name=name, value=value):
                data = self.tiny(); data.arrays[name][:] = value
                self.create_input(data, expected=INVALID)
        for cols in ([1, 0], [0, 0]):
            data = OwnedInput([[1., 2.]], [1.], [0., 0.], [1], [1])
            data.arrays['col_index'][:] = cols
            self.create_input(data, expected=INVALID)
        data = OwnedInput(np.zeros((2, 1)), [0., 0.], [0.], [3], [2])
        self.create_input(data, expected=INVALID)

    def test_handle_and_result_lifecycle_with_output_alias_guards(self):
        data = self.tiny(); h = self.create_input(data)
        x, z, p = np.zeros(1), np.zeros(1), progress()
        try:
            self.assertEqual(self.result(h, x.ctypes.data_as(PD), 1, z.ctypes.data_as(PD), 1, C.byref(p)), STATE)
            self.assertEqual(self.start(h), OK)
            self.await_terminal(h)
            self.assertEqual(self.start(h), STATE)
            self.assertEqual(self.result(h, x.ctypes.data_as(PD), 0, z.ctypes.data_as(PD), 1, C.byref(p)), BUFFER)
            self.assertEqual(self.result(h, None, 1, z.ctypes.data_as(PD), 1, C.byref(p)), INVALID)
            self.assertEqual(self.result(h, x.ctypes.data_as(PD), 1, x.ctypes.data_as(PD), 1, C.byref(p)), INVALID)
            alias = C.cast(C.byref(p), PD)
            self.assertEqual(self.result(h, alias, 1, z.ctypes.data_as(PD), 1, C.byref(p)), INVALID)
            bad = progress(); bad.abi_version = 2
            self.assertEqual(self.poll(h, C.byref(bad)), INVALID)
            self.assertEqual(self.cancel(h), OK)
            polled, returned = progress(), progress()
            self.assertEqual(self.poll(h, C.byref(polled)), OK)
            self.assertEqual(self.result(h, x.ctypes.data_as(PD), 1,
                z.ctypes.data_as(PD), 1, C.byref(returned)), OK)
            self.assertEqual(polled.cancel_requested, returned.cancel_requested)
            self.assertEqual(returned.cancel_requested, 1)
        finally:
            self.cleanup(h)
        for fn in (self.start, self.cancel, self.destroy):
            self.assertEqual(fn(h), HANDLE)
            self.assertEqual(fn(0), HANDLE)
        self.assertEqual(self.poll(h, C.byref(progress())), HANDLE)

    def test_prestart_cancel_and_running_destroy_are_nonblocking(self):
        h = self.create_input(self.tiny())
        try:
            self.assertEqual(self.cancel(h), OK)
            self.assertEqual(self.start(h), OK)
            p = self.await_terminal(h)
            self.assertEqual((p.state, p.cancel_requested), (CANCELLED, 1))
            x, z = np.zeros(1), np.zeros(1)
            # Cancel before construction has no field; status is never a proof.
            self.assertEqual(self.result(h, x.ctypes.data_as(PD), 1,
                z.ctypes.data_as(PD), 1, C.byref(progress())), STATE)
        finally:
            self.cleanup(h)
        h = self.create_input(self.tiny())
        self.assertEqual(self.start(h), OK)
        t = time.monotonic(); rc = self.destroy(h)
        self.assertLess(time.monotonic() - t, 1.)
        self.assertIn(rc, (OK, BUSY))  # A tiny solve can finish before destroy wins the race.
        if rc == BUSY:
            self.assertIn(self.await_terminal(h).state, (DONE, CANCELLED))
            self.cleanup(h)

    def test_registry_limit_is_rejected_and_all_created_handles_release(self):
        handles = []
        try:
            for _ in range(64):
                handles.append(self.create_input(self.tiny()))
            self.create_input(self.tiny(), expected=LIMIT)
        finally:
            for h in handles:
                self.cleanup(h)
        h = self.create_input(self.tiny()); self.cleanup(h)

    def audit_fixture(self, name, prepared):
        data = OwnedInput.original(prepared)
        x, z, p = self.run_input(data)
        xp, zp = self.python_solution(data)
        audit, reference = core.evaluate(prepared, x, z), core.evaluate(prepared, xp, zp)
        self.assertTrue(audit['physical_pass'], audit)
        self.assertTrue(audit['witness'], audit)
        self.assertTrue(reference['physical_pass'], reference)
        self.assertTrue(reference['witness'], reference)
        self.assertLess(abs(audit['objective'] - reference['objective']),
            1e-7 * max(1., abs(reference['objective'])))
        self.records.append({'name': name, 'mode': prepared.mode, 'fs': prepared.fs, 'q': 8,
            'elements': len(prepared.m.tri), 'n': len(x), 'm': len(z),
            'original_nonzero_h': bool(np.any(prepared.original[3])),
            'native_status_metadata': p.solver_status, 'native_original_audit': audit,
            'python_original_audit': reference, 'publishes_bound': False})
        return data, x, z

    def test_real_cohesive_original_q8_lower_and_upper(self):
        for mode, fs in [('lower', .01), ('upper', 100.)]:
            with self.subTest(mode=mode):
                prepared = core.Prepared(physical_triangle(), fs, mode, 8, reduction='none')
                if mode == 'lower':
                    self.assertTrue(np.any(prepared.original[3]), 'exercise affine cohesion offset h')
                self.audit_fixture('real_cohesive_triangle_' + mode, prepared)

    def test_original_eight_triangle_c0_wedge_q8(self):
        _, record = driver.load_model('D02A')
        mesh, _, _ = wedge.wedge_mesh(record, .01)
        prepared = core.Prepared(mesh, 30 / 19.99 + .001, 'upper', 8, reduction='none')
        self.assertEqual(len(mesh.tri), 8)
        self.audit_fixture('original_D02A_eight_triangle_wedge', prepared)

    def test_modified_qp_dual_is_not_an_original_lp_guide(self):
        prepared = core.Prepared(physical_triangle(), .01, 'lower', 8, reduction='none')
        original = OwnedInput.original(prepared)
        x, _, _ = self.run_input(original)
        alpha = .01
        modified = OwnedInput.original(prepared, p=np.full(len(x), alpha), q=-alpha * x)
        candidate, auxiliary_z, _ = self.run_input(modified)
        physical = core.evaluate(prepared, candidate)
        self.assertTrue(physical['physical_pass'], physical)
        rejected = audit_original_dual(prepared, auxiliary_z)
        self.assertFalse(rejected['accepted'], 'QP stationarity cannot be used as original LP stationarity')
        self.assertGreater(rejected['dual_stationarity'], .1)
        self.records.append({'name': 'modified_proximal_qp_original_audit',
            'physical': physical, 'auxiliary_dual_original_gate': rejected,
            'auxiliary_dual_adopted': False, 'publishes_bound': False})

    def equality_input(self, A, b, x0=None):
        data = OwnedInput(A, b, np.zeros(sparse.csr_matrix(A).shape[1]), [], [])
        data.arrays['x0'] = np.zeros(data.problem.n) if x0 is None else np.asarray(x0, dtype=float).copy()
        p = EqualityProblem(struct_size=C.sizeof(EqualityProblem), abi_version=1,
            n=data.problem.n, m=data.problem.m, nnz=data.problem.nnz)
        for name in ('row_ptr', 'col_index', 'A_values', 'b', 'x0'):
            a = data.arrays[name]
            setattr(p, name, a.ctypes.data_as(PI if a.dtype == np.int32 else PD))
        return data, p

    def run_equalities(self, A, b, x0=None, budget=100, atol=1e-15, btol=1e-15, conlim=1e8):
        data, p = self.equality_input(A, b, x0)
        originals = {k: a.tobytes() for k, a in data.arrays.items()}
        x = np.full(p.n, math.nan)
        report = EqualityReport(struct_size=C.sizeof(EqualityReport), abi_version=1)
        rc = self.equalities(C.byref(p), budget, atol, btol, conlim,
            x.ctypes.data_as(PD), len(x), C.byref(report))
        self.assertEqual(rc, OK)
        self.assertEqual(originals, {k: a.tobytes() for k, a in data.arrays.items()})
        self.assertTrue(np.isfinite(x).all())
        return x, report

    def test_equalities_under_over_rank_deficient_and_inconsistent_scipy_parity(self):
        over = np.array([[1., 2., -1.], [0., 2., 3.], [2., 0., 1.],
            [1., -1., 0.], [3., 1., 1.]])
        fixtures = [
            ('overdetermined_consistent', over, over @ np.array([1., 2., 3.]), np.zeros(3)),
            ('underdetermined_nonzero_initial', over.T, np.array([1., 2., 3.]), np.arange(5) / 7),
            ('rank_deficient_inconsistent', np.array([[1., 1.], [2., 2.], [3., 3.]]),
                np.array([1., 2., 4.]), np.zeros(2)),
            ('overdetermined_inconsistent', over, np.arange(5.), np.zeros(3)),
            ('zero_matrix_nonzero_rhs', np.zeros((3, 2)), np.ones(3), np.array([2., 3.])),
            ('empty_rows', np.zeros((0, 2)), np.zeros(0), np.array([2., 3.]))]
        self.assertEqual([C.sizeof(EqualityProblem), C.sizeof(EqualityReport)], [64, 56])
        for name, A, b, x0 in fixtures:
            with self.subTest(name=name):
                x, native = self.run_equalities(A, b, x0)
                reference = core.lsmr(sparse.csr_matrix(A), b, atol=1e-15, btol=1e-15,
                    conlim=1e8, maxiter=100, x0=x0)
                np.testing.assert_allclose(x, reference[0], atol=1e-11, rtol=1e-11)
                self.assertLessEqual(native.iterations, 100)
                if native.istop in (0, 2):
                    self.assertEqual(native.istop, reference[1])
                # A roundoff-sized extra bidiagonalization can change norm(A)
                # and condition *estimates*. Compare the solution and true
                # residuals; iteration/condition diagnostics do not certify it.
                np.testing.assert_allclose([native.normr, native.normar, native.normx],
                    [reference[3], reference[4], reference[7]], atol=1e-10, rtol=1e-9)
                self.assertAlmostEqual(native.normr, float(np.linalg.norm(b - A @ x)), places=10)
                self.assertLess(float(np.linalg.norm(A.T @ (b - A @ x))), 1e-10)
                self.assertTrue(math.isfinite(native.norma) and math.isfinite(native.conda))
                self.records.append({'name': 'lsmr_' + name,
                    'max_x_difference': float(np.max(abs(x - reference[0]), initial=0.)),
                    'native_istop_metadata': native.istop, 'native_iterations': native.iterations,
                    'publishes_bound': False})

    def test_equalities_staged_initial_correction_preserves_original_rhs(self):
        rng = np.random.default_rng(816)
        A = sparse.csr_matrix(rng.normal(size=(23, 19)))
        b = rng.normal(size=23)
        x, xp = np.zeros(19), np.zeros(19)
        actual_iterations = 0
        for budget in (2, 5, 10):
            stage = budget - actual_iterations
            x, native = self.run_equalities(A, b, x, budget=stage)
            reference = core.lsmr(A, b, atol=1e-15, btol=1e-15, maxiter=stage, x0=xp)
            xp = reference[0]
            np.testing.assert_allclose(x, xp, atol=1e-10, rtol=1e-10)
            self.assertEqual(native.iterations, reference[2])
            actual_iterations += native.iterations

    def test_equalities_real_raw_physical_rows_recover_then_q8_full_gate(self):
        prepared = core.Prepared(physical_triangle(), .01, 'lower', 8, reduction='none')
        x, _, _ = self.run_input(OwnedInput.original(prepared))
        source = x * .8
        source[0] += 1e-4
        original_source = source.copy()
        before = core.evaluate(prepared, source)
        self.assertFalse(before['physical_pass'])
        E, b, *_ = prepared.original
        # Equivalent to VBA raw original compiled rows: physical row_norm * E.
        matrix = sparse.diags(prepared.original_row_norm) @ E
        rhs = prepared.original_row_norm * (b - E @ source)
        correction, native = self.run_equalities(matrix, rhs, budget=100)
        reference = core.lsmr(matrix, rhs, atol=1e-15, btol=1e-15, maxiter=100,
            x0=np.zeros(len(source)))
        np.testing.assert_allclose(correction, reference[0], atol=1e-11, rtol=1e-8)
        np.testing.assert_array_equal(source, original_source)
        audit = core.evaluate(prepared, source + correction)
        self.assertTrue(audit['physical_pass'], audit)
        self.assertTrue(audit['witness'], audit)
        self.records.append({'name': 'raw_lower_physical_equality_correction',
            'before_original_gate': before, 'after_original_gate': audit,
            'native_istop_metadata': native.istop, 'publishes_bound': False})

    def test_equalities_rejects_bad_inputs_aliases_and_nonfinite_outputs(self):
        for field, value, expected in [('abi_version', 2, INVALID), ('reserved', 1, INVALID),
            ('n', 0, LIMIT), ('n', 2_000_001, LIMIT), ('m', 8_000_001, LIMIT),
            ('nnz', 50_000_001, LIMIT)]:
            data, p = self.equality_input([[1.]], [1.])
            setattr(p, field, value)
            x = np.full(1, 123.)
            report = EqualityReport(struct_size=56, abi_version=1)
            self.assertEqual(self.equalities(C.byref(p), 100, 1e-15, 1e-15, 1e8,
                x.ctypes.data_as(PD), 1, C.byref(report)), expected)
            self.assertEqual(x[0], 123.)
        for field in ('A_values', 'b', 'x0'):
            data, p = self.equality_input([[1.]], [1.])
            data.arrays[field][0] = math.nan
            self.assertEqual(self.equalities(C.byref(p), 100, 1e-15, 1e-15, 1e8,
                np.zeros(1).ctypes.data_as(PD), 1,
                C.byref(EqualityReport(struct_size=56, abi_version=1))), INVALID)
        for name, values in [('row_ptr', [0, 0]), ('row_ptr', [0, -1]),
            ('col_index', [-1]), ('col_index', [1])]:
            data, p = self.equality_input([[1.]], [1.])
            data.arrays[name][:] = values
            self.assertEqual(self.equalities(C.byref(p), 100, 1e-15, 1e-15, 1e8,
                np.zeros(1).ctypes.data_as(PD), 1,
                C.byref(EqualityReport(struct_size=56, abi_version=1))), INVALID)
        for columns in ([1, 0], [0, 0]):
            data, p = self.equality_input([[1., 2.]], [1.])
            data.arrays['col_index'][:] = columns
            self.assertEqual(self.equalities(C.byref(p), 100, 1e-15, 1e-15, 1e8,
                np.zeros(2).ctypes.data_as(PD), 2,
                C.byref(EqualityReport(struct_size=56, abi_version=1))), INVALID)
        data, p = self.equality_input([[1.]], [1.])
        for budget, atol in [(0, 1e-15), (1001, 1e-15), (100, -1.), (100, math.nan)]:
            self.assertEqual(self.equalities(C.byref(p), budget, atol, 1e-15, 1e8,
                np.zeros(1).ctypes.data_as(PD), 1,
                C.byref(EqualityReport(struct_size=56, abi_version=1))), INVALID)
        report = EqualityReport(struct_size=56, abi_version=1)
        self.assertEqual(self.equalities(C.byref(p), 100, 1e-15, 1e-15, 1e8,
            p.x0, 1, C.byref(report)), INVALID)
        self.assertEqual(self.equalities(C.byref(p), 100, 1e-15, 1e-15, 1e8,
            np.zeros(1).ctypes.data_as(PD), 0, C.byref(report)), BUFFER)
        data, p = self.equality_input([[1e308]], [1.], [1e308])
        out = np.full(1, 123.)
        self.assertEqual(self.equalities(C.byref(p), 100, 1e-15, 1e-15, 1e8,
            out.ctypes.data_as(PD), 1, C.byref(report)), -7)
        self.assertEqual(out[0], 123., 'failed arithmetic must not publish a correction')


if __name__ == '__main__':
    unittest.main()
