"""Compile/CLI and fabricated b=0, h!=0 original lower-coordinate checks."""
from pathlib import Path
import hashlib
import importlib.util
import json
import subprocess
import sys
from types import SimpleNamespace

sys.dont_write_bytecode = True
HERE = Path(__file__).resolve().parent


def run():
    files = sorted(HERE.glob('*.py'))
    for path in files:
        compile(path.read_text(encoding='utf8'), str(path), 'exec')
    dummy = str(HERE / 'not_loaded_seed')
    fresh = str(HERE / 'validation_output_not_created')
    helper = ['--seed', dummy, '--output', fresh, '--guide-fs', '.678378841819045',
              '--seconds', '180', '--regularization', '1e-8', '--validate-args-only']
    close = ['--seed', dummy, '--guide', dummy + '/not_loaded_original_guide.npz',
             '--output', fresh, '--validate-args-only']
    cases = [
        ('solve_lower_scaled_original_helper.py', ['--help'], 0),
        ('close_from_saved_lower_scaled_guide.py', ['--help'], 0),
        ('solve_lower_scaled_original_helper.py', helper, 0),
        ('solve_lower_scaled_original_helper.py', helper + ['--case', 'A', '--equilibrate'], 0),
        ('close_from_saved_lower_scaled_guide.py', close + ['--case', 'N'], 0),
        ('solve_lower_scaled_original_helper.py', helper + ['--seconds', '0'], 2),
        ('solve_lower_scaled_original_helper.py', helper + ['--guide-fs', 'nan'], 2),
        ('solve_lower_scaled_original_helper.py', helper + ['--regularization', '-1'], 2),
        ('solve_lower_scaled_original_helper.py', helper + ['--y-scale-max', 'inf'], 2),
        ('solve_lower_scaled_original_helper.py', helper + ['--y-scale-min', '10', '--y-scale-max', '1'], 2),
        ('close_from_saved_lower_scaled_guide.py', close + ['--tolerance', '.0002'], 2),
        ('close_from_saved_lower_scaled_guide.py', close + ['--seconds', 'inf'], 2),
    ]
    checks = []
    for name, args, expected in cases:
        result = subprocess.run([sys.executable, '-B', '-X', 'utf8', str(HERE / name), *args],
                                text=True, capture_output=True)
        assert result.returncode == expected, (name, args, result.stdout, result.stderr)
        checks.append({'script': name, 'args': args, 'exit_code': result.returncode,
                       'expected_exit_code': expected,
                       'last_output': (result.stdout + result.stderr).strip().splitlines()[-1]})
    import lower_scaled_original_package as package
    import numpy as np
    from scipy import sparse
    fake_core = SimpleNamespace(np=np, sparse=sparse)
    G_array = np.zeros((6, 19))
    G_array[:, :4] = np.array([
        [20., 0., 2., 0.], [0., 4., 0., 0.], [0., 0., 5., 0.],
        [2e5, 0., 3., 0.], [0., 1e-4, 0., 0.], [0., 0., 30., 2.],
    ])
    G = sparse.csc_matrix(G_array)
    b, h, dims = np.zeros(2), np.array([.05, 0., 0., .1, 0., 0.]), (3, 3)
    cost = np.zeros(19)
    cost[-1] = -1.
    Y, Z = np.array([-2., .5]), np.array([5., 3., 4., 6., 1., 2.])
    E_array = np.zeros((2, 19))
    E_array[0, :2] = [1e-14, -2e-15]
    E_array[1] = (np.asarray(G.T @ Z).ravel() - cost - Y[0] * E_array[0]) / Y[1]
    E = sparse.csc_matrix(E_array)
    original = (E, b, G, h, dims, cost)
    prepared = SimpleNamespace(original=original)
    before = package.program_hash(fake_core, original)
    ky, kz, scale_report = package.coordinate_scales(fake_core, prepared, 1e-8, 1e8)
    assert ky[0] == 1e8 and np.all(ky > 0) and not ky.flags.writeable and not kz.flags.writeable
    assert np.array_equal(kz, np.array([.05, .05, .05, 5e-6, 5e-6, 5e-6]))
    new_coordinates = np.r_[Y / ky, Z / kz]
    threshold = -.999998
    matrix, rhs = package.transformed_constraints(fake_core, prepared, ky, kz, threshold)
    values = np.asarray(matrix @ new_coordinates).ravel()
    stationary_residual = values[:len(cost)] - rhs[:len(cost)]
    assert np.max(abs(stationary_residual), initial=0.) <= 1e-9
    assert np.array_equal(rhs[:len(cost)], -cost) and rhs[len(cost) - 1] == 1.
    original_separation = float(np.r_[b, h] @ np.r_[Y, Z])
    transformed_separation = float(values[-1])
    assert np.isclose(original_separation, transformed_separation, rtol=1e-13, atol=1e-13)
    assert original_separation > 0 and transformed_separation <= rhs[-1]
    assert np.array_equal(matrix[-1, :len(b)].toarray(), np.zeros((1, len(b))))
    assert np.allclose(matrix[-1, len(b):].toarray().ravel(), h * kz, rtol=0, atol=0)
    assert not np.allclose(matrix[-1, len(b):].toarray().ravel(), h, rtol=0, atol=0)
    restored = np.r_[ky * new_coordinates[:len(b)], kz * new_coordinates[len(b):]]
    assert np.allclose(restored, np.r_[Y, Z], rtol=1e-14, atol=1e-14)
    cone_checks, offset = [], 0
    for dimension in dims:
        old = Z[offset:offset + dimension]
        new = -values[len(cost) + offset:len(cost) + offset + dimension]
        old_margin = float(old[0] - np.linalg.norm(old[1:]))
        new_margin = float(new[0] - np.linalg.norm(new[1:]))
        assert old_margin >= -1e-12 and new_margin >= -1e-12
        assert np.all(kz[offset:offset + dimension] == kz[offset])
        assert np.isclose(old_margin, kz[offset] * new_margin, rtol=1e-13, atol=1e-12)
        cone_checks.append({'dimension': dimension, 'old_margin': old_margin,
                            'new_margin': new_margin, 'common_positive_scale': float(kz[offset]),
                            'membership_preserved': True})
        offset += dimension
    anisotropic_bad = Z[:3] * np.array([1., 1., 2.])
    assert anisotropic_bad[0] < np.linalg.norm(anisotropic_bad[1:])
    assert package.program_hash(fake_core, original) == before
    assert np.array_equal(b, np.zeros_like(b)) and np.count_nonzero(h) == 2
    assert cost.shape == (19,) and cost[-1] == -1. and np.count_nonzero(cost) == 1
    earlier = Path(r'D:\Work\rpfem-n-helper-review')
    old_lower = json.loads((earlier / 'compile_and_args_validation.json').read_text(encoding='utf8'))
    old_upper = json.loads((earlier / 'upper_scaled_compile_and_args_validation.json').read_text(encoding='utf8'))
    old_lower_sha = {name: hashlib.sha256((earlier / name).read_bytes()).hexdigest()
                     for name in old_lower['final_script_sha256']}
    assert old_lower_sha == old_lower['final_script_sha256']
    assert hashlib.sha256((earlier / 'solve_upper_scaled_original_helper.py').read_bytes()).hexdigest() == old_upper['script_sha256']
    assert 'clarabel' not in sys.modules and 'experiment_core' not in sys.modules
    assert not Path(fresh).exists()
    report = {
        'compile_pass': True, 'cli_checks': checks, 'synthetic_transform_checks_pass': True,
        'synthetic_stationarity_max_residual': float(np.max(abs(stationary_residual), initial=0.)),
        'synthetic_b_exactly_zero': True, 'synthetic_h_nonzero_count': int(np.count_nonzero(h)),
        'synthetic_original_separation': original_separation,
        'synthetic_transformed_separation': transformed_separation,
        'synthetic_full_h_times_kz_preserved': True, 'synthetic_free_load_cost_last': float(cost[-1]),
        'synthetic_lambda_stationary_rhs': float(rhs[len(cost) - 1]),
        'synthetic_cone_checks': cone_checks, 'synthetic_scale_report': scale_report,
        'anisotropic_cone_scaling_counterexample_detected': True,
        'synthetic_original_program_sha256_unchanged': True,
        'synthetic_only_not_a_physical_audit': True,
        'no_source_model_inputs_loaded': True, 'no_repository_modules_loaded': True,
        'no_solver_or_root_close_started': True, 'no_output_directory_created': True,
        'old_lower3_and_upper_prototype_byte_identical': True,
        'new_package_sha256': {path.name: hashlib.sha256(path.read_bytes()).hexdigest() for path in files},
        'limit': 'Finite positive coordinates preserve the feasible set; raw restored original dual acceptance can still fail.',
    }
    (HERE / 'compile_cli_synthetic_validation.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf8')
    print(json.dumps({key: report[key] for key in ('compile_pass', 'synthetic_transform_checks_pass',
        'synthetic_stationarity_max_residual', 'synthetic_original_separation',
        'synthetic_transformed_separation', 'synthetic_full_h_times_kz_preserved',
        'no_source_model_inputs_loaded', 'no_solver_or_root_close_started',
        'old_lower3_and_upper_prototype_byte_identical', 'new_package_sha256')}, indent=2))
    print('CLI_CHECKS_PASSED', len(checks))


if __name__ == '__main__':
    run()
