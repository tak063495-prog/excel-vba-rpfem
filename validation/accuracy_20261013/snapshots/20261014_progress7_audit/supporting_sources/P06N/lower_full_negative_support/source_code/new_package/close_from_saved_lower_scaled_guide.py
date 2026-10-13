"""Cached own modern lower close from an exact frozen scaled-helper packet."""
import argparse
import json
from pathlib import Path
import time
import lower_scaled_original_package as package


def parser():
    p = argparse.ArgumentParser(description=__doc__)
    p.add_argument('--case', default='P06N', type=package.case_name)
    p.add_argument('--seed', type=Path, required=True)
    p.add_argument('--guide', type=Path, required=True)
    p.add_argument('--output', type=Path, required=True)
    p.add_argument('--seconds', type=float, default=120.)
    p.add_argument('--tolerance', type=float, default=1e-4)
    p.add_argument('--validate-args-only', action='store_true')
    return p


def validate(args):
    package.validate_paths(args)
    package.positive(args.seconds, 'Seconds')
    package.positive(args.tolerance, 'Tolerance')
    package.require(args.tolerance <= 1e-4, 'Tolerance must be at most 1e-4')
    package.require(not args.guide.resolve().is_relative_to(args.output.resolve()), 'Guide must be outside output')


def root_guard(root, best, identity):
    width, bracket = root.get('bracket_width'), root.get('bracket')
    required = {key: identity[key] for key in ('case', 'mode', 'q', 'mesh_id', 'physical_model_id', 'policy')}
    return bool(best and best.root_complete and root.get('root_complete')
        and all(root.get(key) == value for key, value in required.items())
        and root.get('bracket_signs') == [1, -1]
        and isinstance(width, (float, int)) and package.math.isfinite(width) and 0 < width <= 1e-4
        and isinstance(bracket, list) and len(bracket) == 2
        and bracket[0] == identity['fs'] == root.get('endpoint') == best.fs
        and bracket[1] > bracket[0] and bracket[1] - bracket[0] == width
        and root.get('opposite_numerical_endpoint_is_continuum_bound') is False)


def run(args):
    validate(args)
    if args.validate_args_only:
        print('LOWER_SCALED_CLOSE_ARGUMENTS_VALID_NO_INPUTS_READ_NO_CLOSE', flush=True)
        return None
    started = time.perf_counter()
    frozen = package.Frozen(args, args.guide)
    core, driver, original_dual = frozen.core, frozen.driver, frozen.original_dual
    try:
        mesh, field, source_fs, identity, witness, source, source_hash, source_layout = frozen.load_source()
        with core.np.load(frozen.guide_copy, allow_pickle=False) as data:
            for key in ('case', 'mode', 'q', 'mesh_id', 'physical_model_id', 'policy'):
                package.scalar(data, key, identity[key])
            package.scalar(data, 'source_fs', source_fs)
            package.require(data['fs'].shape == (), 'Guide Fs must be scalar')
            guide_fs = float(data['fs'].item())
            package.positive(guide_fs, 'Guide Fs')
            for key, expected in frozen.provenance_packets().items():
                package.require(key in data.files and data[key].shape == (), 'Missing scalar frozen provenance packet: ' + key)
                package.require(json.loads(str(data[key].item())) == expected, 'Cached helper frozen provenance differs: ' + key)
            # Stored acceptance/root flags are ignored. Only the restored full
            # original vector is read; the unchanged original gate decides.
            dual = data['dual'].copy()
        package.require(guide_fs > source_fs, 'Original lower guide Fs must exceed source Fs')
        package.require(guide_fs - source_fs <= args.tolerance,
                        'Cached close requires an already narrow bracket; generate a guide within tolerance before closing')
        guide, guide_hash, guide_layout = package.original_lower(core, original_dual, mesh, guide_fs)
        package.require(dual.shape == (guide_layout['full_original_dual_shape'][0],)
                        and core.np.isfinite(dual).all(), 'Finite full original dual layout required')
        guide_audit = original_dual.audit_original_dual(guide, dual)
        package.require(guide_audit['accepted'] and guide_audit.get('search_sign') == -1, 'Fresh original lower dual gate rejected guide')
        package.require(guide_hash == package.program_hash(core, guide.original), 'Original guide program changed during audit')
        before_search = frozen.check('after_fresh_source_and_guide_gates_before_cached_search')
        initial = driver.Certificate('lower', source_fs, mesh, field, witness,
                                     'Fresh original q8 lower witness and exact scaled-helper restored original guide', False)
        original_guide = dict(identity, fs=guide_fs, dual=dual)
        label = args.case + '_lower_scaled_original_guide_cached_close'
        staging = frozen.output / 'internal_driver_search' / args.case
        best, root = driver.search(mesh, 'lower', (source_fs, guide_fs), staging,
                                  label, args.tolerance, args.seconds, initial, [original_guide])
        package.require(source_hash == package.program_hash(core, source.original)
                        and guide_hash == package.program_hash(core, guide.original), 'Original source/guide programs changed during cached search')
        after_search = frozen.check('after_driver_own_modern_cached_search')
        package.require(root_guard(root, best, identity), 'Driver own modern root does not meet exact source endpoint/identity/sign/width guards')
        package.require(root['trials'] == 2, 'Cached close unexpectedly evaluated additional strengths')
        final_source, final_source_hash, final_source_layout = package.original_lower(core, original_dual, mesh, source_fs)
        final_guide, final_guide_hash, final_guide_layout = package.original_lower(core, original_dual, mesh, guide_fs)
        package.require(final_source_hash == source_hash and final_guide_hash == guide_hash
                        and final_source_layout == source_layout and final_guide_layout == guide_layout,
                        'Fresh original source/guide programs differ after cached close')
        final_witness = core.evaluate(final_source, best.field, None)
        final_guide_audit = original_dual.audit_original_dual(final_guide, dual)
        package.require(best.field.shape == (18 * len(mesh.tri) + 1,)
                        and core.np.array_equal(best.field, field)
                        and final_witness['physical_pass'] and final_witness['witness']
                        and final_guide_audit['accepted'], 'Final fresh original source/guide raw gates failed')
        package.require(final_source_hash == package.program_hash(core, final_source.original)
                        and final_guide_hash == package.program_hash(core, final_guide.original), 'Final original programs changed during audits')
        final_witness.update(identity, discrete_root_search=root, source_flags_ignored=True,
                             origin='final_fresh_original_lower_witness_after_own_modern_cached_close')
        best.result, best.root_complete = final_witness, True
        best.publish(staging)
        final_freeze = frozen.check('before_final_witness_publication')
        folder = frozen.output / args.case
        folder.mkdir()
        published = best.publish(folder)
        core.write(folder / 'summary.json', {'case': args.case, 'lower': published,
            'lower_only': True, 'root': root, 'one_percent_target_not_evaluated': True})
        core.write(folder / 'import_audit.json', {'case': args.case, 'certificates': {'lower': published}})
        report = {**identity, 'source_fs': source_fs, 'guide_fs': guide_fs,
            'root': root, 'root_complete': True, 'source_flags_ignored': True,
            'guide_flags_ignored': True, 'auxiliary_solution_z_read_or_used': False,
            'source_original_audit': witness, 'fresh_original_dual_audit': guide_audit,
            'final_source_original_audit': final_witness, 'final_original_dual_audit': final_guide_audit,
            'source_original_layout': source_layout, 'guide_original_layout': guide_layout,
            'source_original_program_sha256_before': source_hash,
            'source_original_program_sha256_after': final_source_hash,
            'guide_original_program_sha256_before': guide_hash,
            'guide_original_program_sha256_after': final_guide_hash,
            'sourcefreeze_before': frozen.before, 'sourcefreeze_before_search': before_search,
            'sourcefreeze_after_search': after_search, 'sourcefreeze_final': final_freeze,
            'cached_endpoints_only': True, 'seconds': time.perf_counter() - started}
        core.write(frozen.output / 'closure_audit.json', report)
        frozen.check('after_final_witness_publication')
        print('SAVED_SCALED_ORIGINAL_LOWER_CACHED_CLOSE', best.fs, True,
              root['bracket'], root['bracket_width'], flush=True)
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
