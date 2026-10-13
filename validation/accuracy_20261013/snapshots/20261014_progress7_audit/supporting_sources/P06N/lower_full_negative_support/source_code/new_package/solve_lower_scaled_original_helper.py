"""Guide-only original lower dual feasibility in positive coordinates."""
import argparse
import json
from pathlib import Path
import time
import lower_scaled_original_package as package


def parser():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--case', default='P06N', type=package.case_name)
    p.add_argument('--seed', type=Path, required=True)
    p.add_argument('--output', type=Path, required=True)
    p.add_argument('--guide-fs', type=float, required=True)
    p.add_argument('--seconds', type=float, default=180.)
    p.add_argument('--regularization', type=float, default=1e-8)
    p.add_argument('--equilibrate', action='store_true')
    p.add_argument('--y-scale-min', type=float, default=1e-8)
    p.add_argument('--y-scale-max', type=float, default=1e8)
    p.add_argument('--validate-args-only', action='store_true')
    return p


def validate(args):
    package.validate_paths(args)
    package.positive(args.guide_fs, 'Guide Fs')
    package.positive(args.seconds, 'Seconds')
    package.positive(args.y_scale_min, 'Y scale minimum')
    package.positive(args.y_scale_max, 'Y scale maximum')
    package.require(args.y_scale_min <= args.y_scale_max, 'Y scale minimum must not exceed maximum')
    package.require(package.math.isfinite(args.regularization) and args.regularization >= 0,
                    'Regularization must be finite and nonnegative')


def run(args):
    validate(args)
    if args.validate_args_only:
        print('LOWER_SCALED_HELPER_ARGUMENTS_VALID_NO_INPUTS_READ_NO_SOLVE', flush=True)
        return None
    started = time.perf_counter()
    frozen = package.Frozen(args)
    core, original_dual = frozen.core, frozen.original_dual
    try:
        mesh, field, source_fs, identity, witness, source, source_hash, source_layout = frozen.load_source()
        package.require(args.guide_fs > source_fs, 'Lower guide Fs must be strictly above the source witness Fs')
        guide, guide_hash, guide_layout = package.original_lower(core, original_dual, mesh, args.guide_fs)
        before_helper = frozen.check('before_scaled_original_helper')
        full, ky, kz, scale_report, evidence = package.solve_scaled_helper(core, original_dual, guide, args)
        package.require(source_hash == package.program_hash(core, source.original)
                        and guide_hash == package.program_hash(core, guide.original), 'Original physical programs changed during helper')
        after_helper = frozen.check('after_scaled_helper_before_fresh_raw_audit')
        fresh, fresh_hash, fresh_layout = package.original_lower(core, original_dual, mesh, args.guide_fs)
        package.require(fresh_hash == guide_hash and fresh_layout == guide_layout, 'Fresh original guide program differs')
        raw_audit = original_dual.audit_original_dual(fresh, full)
        package.require(fresh_hash == package.program_hash(core, fresh.original), 'Fresh physical program changed during raw dual audit')
        after_audit = frozen.check('after_fresh_original_restored_dual_audit')
        scales_path = frozen.output / 'coordinate_scales.npz'
        core.np.savez_compressed(scales_path, ky=ky, kz=kz, dims=core.np.asarray(guide.original[4], dtype=core.np.int64))
        full_path = frozen.output / 'original_helper_full_dual.npz'
        packets = {key: json.dumps(value, sort_keys=True) for key, value in frozen.provenance_packets().items()}
        core.np.savez_compressed(full_path, dual=full, **dict(identity, fs=args.guide_fs), **packets,
            source_fs=source_fs, guide_only=True, root_complete=False, publishes_primal_bound=False,
            restored_original_coordinates=True, accepted=bool(raw_audit['accepted']))
        report = {**identity, 'source_fs': source_fs, 'guide_fs': args.guide_fs,
            'guide_identity': dict(identity, fs=args.guide_fs), 'guide_only': True,
            'scope': 'discrete_root_search_only', 'root_complete': False,
            'publishes_primal_bound': False, 'physical_field_published': False,
            'source_flags_ignored': True, 'auxiliary_solution_z_read_or_used': False,
            'original_full_vector_source': 'K times helper solution.x; restored original (Y,Z)',
            'source_original_audit': witness, 'helper_evidence': evidence,
            'fresh_original_dual_audit': raw_audit, 'coordinate_scales': scale_report,
            'original_full_dual_sha256': package.sha(full_path), 'scale_artifact_sha256': package.sha(scales_path),
            'source_original_layout': source_layout, 'guide_original_layout': guide_layout,
            'source_original_program_sha256_before': source_hash,
            'source_original_program_sha256_after': package.program_hash(core, source.original),
            'guide_original_program_sha256_before': guide_hash,
            'guide_original_program_sha256_after': package.program_hash(core, guide.original),
            'fresh_guide_original_program_sha256': fresh_hash,
            'sourcefreeze_before': frozen.before, 'sourcefreeze_before_helper': before_helper,
            'sourcefreeze_after_helper': after_helper, 'sourcefreeze_after_fresh_audit': after_audit,
            'sourcefreeze_final': frozen.check('before_guide_only_publication'),
            'seconds': time.perf_counter() - started}
        core.write(frozen.output / 'helper_audit.json', report)
        core.write(frozen.output / 'closure_audit.json', report)
        frozen.check('after_guide_only_publication')
        print('LOWER_SCALED_ORIGINAL_GUIDE_ONLY', source_fs, args.guide_fs,
              raw_audit['accepted'], False, raw_audit, flush=True)
        return report
    except Exception as error:
        package.reject(frozen.output, f'{type(error).__name__}: {error}')
        raise


def main():
    p = parser()
    try:
        return run(p.parse_args())
    except ValueError as error:
        p.error(str(error))


if __name__ == '__main__':
    main()
