"""Verify a posthoc pair against real component starts; never label it cold.

Historical geometry is allowed only in this explicitly selected evidence kind.
The original physical audit is separate and must have already passed.
"""
from pathlib import Path
from evidence_provenance import (REPO, CATALOG, read, require, sha, canonical,
                                 inside, hashes, finite_number, _root, _model_id)


def checked_map(value, repo, label):
    result = hashes(value, repo, label)
    for path, digest in result.items():
        require(sha(path.read_bytes()) == digest, label + ': file changed: ' + str(path))
    return result


def checked_reference(reference, repo):
    path = inside(reference['path'], repo)
    require(sha(path.read_bytes()) == reference['sha256'], 'Referenced component file changed: ' + str(path))
    return path


def checked_unowned_repair(component, repo, case, input_path, record, policy,
                          supervised_sources, supervised_artifacts):
    """Check the explicitly unowned no-optimizer correction, never a cold run.

    Source/input maps are the actual producer maps. Missing PID/process timing
    stays missing; an INPUTS_26 map not read by the producer is not invented.
    """
    require(case == 'P06N' and component.get('role') == 'upper_adopted_linear_repair_closure' and
            component.get('adopted') is True and 'ownership_file' in component and component['ownership_file'] is None,
            'Only the explicitly identified final N repair may lack ownership')
    require(all(k in component and component[k] is None for k in ('started_epoch', 'finished_epoch', 'process_wall_seconds')),
            'Unsupervised repair must not invent process timing')
    folder = inside(component['folder'], repo)
    execution = component.get('execution_record', {})
    ledger_path = checked_reference(execution, repo)
    ledger = read(ledger_path)
    require({k:v for k,v in execution.items() if k not in ('path', 'sha256')} == ledger,
            'Repair execution record differs from its actual posthoc ledger')
    require(ledger.get('kind') == 'captured_exec_without_pid_record' and ledger.get('posthoc_record') is True and
            ledger.get('observed_returncode') == component.get('os_returncode') == 0 and
            component.get('audit_returncode') == 0 and ledger.get('optimizer_calls') == 0,
            'Repair is not an explicit captured exit0, no-optimizer record')
    require(all(k in ledger and ledger[k] is None for k in ('os_pid', 'os_birth_time', 'started_epoch',
                'finished_epoch', 'process_wall_seconds')), 'Repair ledger invented PID or process timing')
    producer_path = checked_reference(component['summary'], repo)
    require(producer_path == folder/'summary.json' and
            checked_reference(ledger['source_summary'], repo) == producer_path,
            'Repair producer summary points elsewhere')
    producer = read(producer_path)
    require(producer.get('case') == case and producer.get('mode') == 'upper' and producer.get('q') == 8 and
            producer.get('root_complete') is True and producer.get('status') == 'OWN_ROOT_COMPLETED' and
            producer.get('source_unchanged') is component.get('source_unchanged') is True,
            'Repair producer identity, completion or source statement differs')
    scope = producer.get('execution_scope', {})
    require(scope.get('kind') == 'current_session_side_root_continuation' and
            scope.get('historical_field_imported') is scope.get('historical_fs_imported') is False and
            scope.get('optimizer_calls') == 0, 'Repair scope is not a current-session no-optimizer continuation')
    command = ledger.get('command')
    require(isinstance(command, list) and len(command) >= 4 and command[1:3] == ['-X', 'utf8'],
            'Missing captured repair command')
    wrapper = inside(command[3], repo)
    require(wrapper.name == 'repair_original_dual.py', 'Different unowned numerical wrapper')
    def argument(flag):
        require(command.count(flag) == 1 and command.index(flag)+1 < len(command),
                'Missing or repeated repair command argument: '+flag)
        return command[command.index(flag)+1]
    require(argument('--case') == case and inside(argument('--output'), repo) == folder,
            'Captured repair command points at another output/case')
    source, source_dual = inside(argument('--source'), repo), inside(argument('--dual'), repo)
    require(inside(producer['source'], repo) == source and inside(producer['source_dual'], repo) == source_dual,
            'Repair source differs from captured command')
    start = checked_map(component['component_start_source_sha256'], repo, 'Repair start source')
    require(start == hashes(component['component_end_source_sha256'], repo, 'Repair end source') ==
            hashes(producer['source_sha256_before'], repo, 'Actual repair before') ==
            hashes(producer['source_sha256_after'], repo, 'Actual repair after') ==
            hashes(ledger['component_start_source_sha256'], repo, 'Ledger repair before') ==
            hashes(ledger['component_end_source_sha256'], repo, 'Ledger repair after'),
            'Actual repair start/end source records differ')
    drivers = [p for p in start if p.name == 'accuracy_driver.py' and p.parent.name == 'python']
    require(len(drivers) == 1, 'One numerical tree per repair is required')
    python = drivers[0].parent
    required = {p.resolve() for p in python.rglob('*.py') if '__pycache__' not in p.parts}
    require(required | {wrapper} <= start.keys(), 'Repair inventory omits numerical code or executed wrapper')
    require(required <= supervised_sources.keys() and all(supervised_sources[p] == start[p] for p in required),
            'Repair complete numerical tree/model builder lacks a matching real supervised freeze')
    model = _model_id(record, python/'benchmark_models.py')
    inputs = checked_map(component['input_provenance'], repo, 'Repair input')
    require(inputs == hashes(producer['input_sha256'], repo, 'Actual repair inputs') ==
            hashes(ledger['input_sha256'], repo, 'Ledger repair inputs') and input_path in inputs,
            'Repair canonical input/provenance differs from actual maps')
    require({source/'best_upper_certificate.json', source/'best_upper_mesh.npz',
             source/'best_upper_field.npz', source_dual} <= inputs.keys(), 'Repair omitted source field/mesh/dual')
    require(source_dual in supervised_artifacts and inputs[source_dual] == supervised_artifacts[source_dual],
            'Repair original dual was not a frozen supervised helper vector')
    require(all(inputs[source/name] in supervised_artifacts.values()
                for name in ('best_upper_mesh.npz', 'best_upper_field.npz')),
            'Repair incoming primal bytes do not match a real supervised component')
    elapsed = producer.get('seconds')
    require(finite_number(elapsed) and elapsed > 0 and elapsed == component.get('wrapper_seconds') == ledger.get('wrapper_seconds'),
            'Repair wrapper elapsed differs from its actual record')
    artifacts = component['component_artifacts_sha256']
    require(bool(artifacts), 'No repair evidence bytes recorded')
    for name, digest in artifacts.items():
        path = (folder/name).resolve()
        require(path.is_relative_to(folder) and sha(path.read_bytes()) == digest,
                'Repair artifact changed or escaped its folder')
    certificate = read(folder/'best_upper_certificate.json')
    identity = _root(certificate, case, 'upper', policy, model)
    require(producer.get('upper') == certificate and producer.get('certificates', {}).get('upper') == certificate,
            'Repair final certificate differs from actual producer summary')
    primal_path = checked_reference(component['independent_primal_audit'], repo)
    primal = read(primal_path)
    require(primal.get('status') == 'VERIFIED_OWN_ROOT' and primal.get('case') == case and
            primal.get('mode') == 'upper' and primal.get('physical_pass') is primal.get('witness') is True and
            primal.get('root_complete') is True and primal.get('fs') == identity['fs'] and
            primal.get('audit', {}).get('discrete_root_search') == certificate['audit']['discrete_root_search'],
            'Independent repair primal/root audit differs')
    dual_path = checked_reference(component['independent_dual_audit'], repo)
    replay = read(dual_path)
    context = dict(case=case, mode='upper', q=8, mesh_id=identity['mesh_id'], physical_model_id=model, policy=policy)
    require(replay.get('status') == 'VERIFIED_ORIGINAL_DUAL' and replay.get('optimizer_calls') == 0 and
            replay.get('source_unchanged') is True and replay.get('context') == context and
            replay.get('fs') == certificate['audit']['discrete_root_search']['bracket'][0],
            'Independent repaired dual context/Fs differs from opposite root endpoint')
    audit = replay.get('audit', {})
    require(audit.get('accepted') is True and audit.get('search_sign') == 1 and
            audit.get('scope') == 'discrete_root_search_only' and audit.get('publishes_primal_bound') is False and
            all(finite_number(audit.get(k)) for k in ('dual_stationarity', 'dual_cone_error', 'dual_objective', 'threshold')) and
            0 <= audit['dual_stationarity'] <= 1e-7 and 0 <= audit['dual_cone_error'] <= 1e-8 and
            audit['threshold'] == 1.000003 and audit['dual_objective'] >= audit['threshold'] and
            producer.get('linear_repair', {}).get('after') == audit and
            producer['linear_repair'].get('z_bitwise_unchanged') is True and
            producer['linear_repair'].get('accepted') is True and producer['linear_repair'].get('optimizer_calls') == 0,
            'Independent repaired original dual raw gate fails or differs from producer')
    replay_inputs = checked_map(replay['code_and_input_sha256_before'], repo, 'Dual replay input')
    require(replay_inputs == hashes(replay['code_and_input_sha256_after'], repo, 'Dual replay end') and
            {input_path, folder/'best_upper_mesh.npz', folder/'repaired_original_dual.npz'} <= replay_inputs.keys(),
            'Independent dual replay omitted or changed original vector/model/mesh')
    require(len([p for p in replay_inputs if p.name == 'audit_repaired_dual.py']) == 1,
            'Independent full-vector dual replay omitted its executed audit wrapper')
    replay_sources = checked_map(replay['source_sha256'], repo, 'Dual replay numerical source')
    require(required <= replay_sources.keys() and all(replay_sources[p] == start[p] for p in required),
            'Independent dual replay used another numerical program')
    checked = dict(role=component['role'], adopted=True, process_seconds=None, wrapper_seconds=elapsed,
        process_exit_code=0, execution_method='captured_exec_without_pid_record', supervised_execution_verified=False,
        folder=str(folder.relative_to(repo)).replace('\\', '/'))
    reference = dict(path=str(dual_path.relative_to(repo)).replace('\\', '/'), sha256=sha(dual_path.read_bytes()),
        fs=replay['fs'], audit=audit, full_original_vector_sha256=replay_inputs[folder/'repaired_original_dual.npz'])
    return model, start, checked, reference


def verify_composite_provenance(summary_path, case, *, repo=REPO, audit=None):
    repo = Path(repo).resolve()
    summary_path = inside(summary_path, repo)
    folder = summary_path.parent
    summary = read(summary_path)
    scope = summary.get('execution_scope', {})
    require(summary.get('case') == case and scope.get('kind') == 'known_recipe_geometry_replay',
            'Not an explicitly identified geometry replay')
    require(scope.get('historical_geometry_imported') is True and
            scope.get('historical_field_imported') is False and
            scope.get('historical_fs_imported') is False and
            scope.get('generic_cold_start') is False and
            scope.get('posthoc_independent_original_q8_collection') is True,
            'Geometry replay scope is incomplete or falsely claims a cold run')
    reference = summary['component_manifest']
    manifest_path = (folder/reference['path']).resolve()
    require(manifest_path.is_relative_to(folder) and sha(manifest_path.read_bytes()) == reference['sha256'],
            'Composite manifest changed or escaped its folder')
    manifest = read(manifest_path)
    require(manifest.get('kind') == scope['kind'] and manifest.get('posthoc_composite_manifest') is True and
            manifest.get('not_a_combined_solver_start_manifest') is True,
            'Posthoc collection must not masquerade as a combined solver start')
    reaudit_sources = checked_map(manifest['collector_reaudit_source_sha256'], repo, 'Collector source')
    input_path = repo/CATALOG/(case+'.json')
    record = read(input_path)
    require(record.get('case') == case and record.get('source_case', {}).get('id') == case,
            'Canonical input case differs')
    policy = record['source_case']['policy']
    models = set()
    checked_components = []
    all_sources = dict(reaudit_sources)
    supervised_sources, supervised_artifacts = {}, {}
    independent_dual_audits = []
    components = manifest.get('components', [])
    require(bool(components), 'No real solver components')
    for component in components:
        run = inside(component['run'], repo)
        component_folder = inside(component['folder'], repo)
        require(component_folder.is_relative_to(run), 'Component folder escaped its run')
        if component.get('ownership_file') is None:
            require(not independent_dual_audits, 'Only one explicitly unsupervised final repair is allowed')
            model, start, checked, dual_reference = checked_unowned_repair(component, repo, case,
                input_path, record, policy, supervised_sources, supervised_artifacts)
            models.add(model)
            checked_components.append(checked)
            independent_dual_audits.append(dual_reference)
            for path, digest in start.items():
                require(path not in all_sources or all_sources[path] == digest, 'Source identity conflict')
                all_sources[path] = digest
            continue
        owner_path = checked_reference(component['ownership_file'], repo)
        owner = read(owner_path)
        require(owner_path == run/'OWNERSHIP.json' and owner.get('status') == 'FINISHED',
                'Component supervisor did not finish')
        start = checked_map(component['component_start_source_sha256'], repo, 'Component start source')
        end = hashes(component['component_end_source_sha256'], repo, 'Component end source')
        require(start == end and component.get('source_unchanged') is True and owner.get('source_unchanged') is True and
                start == hashes(owner['source_sha256'], repo, 'Supervisor start') and
                end == hashes(owner['source_sha256_after'], repo, 'Supervisor end'),
                'Actual component start/end source records differ')
        drivers = [p for p in start if p.name == 'accuracy_driver.py' and p.parent.name == 'python']
        require(len(drivers) == 1, 'One implementation tree per real component is required')
        python = drivers[0].parent
        required = {p.resolve() for p in python.rglob('*.py') if '__pycache__' not in p.parts}
        require(required <= start.keys(), 'Component source inventory omits numerical code')
        input_manifest = python.parent/'INPUTS_26.json'
        require(input_manifest in start, 'Component omitted its canonical case manifest')
        entries = read(input_manifest)['cases']
        require(len(entries) == 26 and len({v['case'] for v in entries}) == 26, 'Invalid 26-case manifest')
        matching = [v for v in entries if v['case'] == case]
        require(len(matching) == 1, 'Component manifest case differs')
        entry = matching[0]
        require(inside(entry['input']['path'], repo) == input_path and entry['policy'] == policy and
                entry['input']['sha256_lf'] == sha(input_path.read_bytes().replace(b'\r\n', b'\n')) and
                entry['input']['sha256_canonical_json'] == sha(canonical(record)), 'Canonical input changed')
        models.add(_model_id(record, python/'benchmark_models.py'))
        inputs = checked_map(component['input_provenance'], repo, 'Component input')
        require(input_path in inputs and input_manifest in inputs, 'Component omitted canonical input provenance')
        jobs = [v for v in owner['jobs'].values() if inside(v['folder'], repo) == component_folder]
        require(len(jobs) == 1, 'Real component job missing or duplicated')
        job = jobs[0]
        require(job['case'] == case and job['phase'] == 'FINISHED' and
                job['returncode'] == component['os_returncode'] and
                job['audit_returncode'] == component['audit_returncode'] == 0,
                'Component exit/audit status differs from real supervisor')
        command_script = inside(job['command'][3], repo)
        require(command_script in start, 'Executed component wrapper was not frozen')
        for key in ('started_epoch', 'finished_epoch'):
            require(job[key] == component[key], 'Component process time differs')
        seconds = component['finished_epoch']-component['started_epoch']
        require(seconds > 0 and abs(seconds-component['process_wall_seconds']) < 1e-8,
                'Component elapsed time differs')
        producer_summary_path = checked_reference(component['summary'], repo)
        require(producer_summary_path == component_folder/'summary.json', 'Component summary points elsewhere')
        producer_summary = read(producer_summary_path)
        require(producer_summary.get('case') == case and producer_summary.get('source_unchanged') is True,
                'Component producer identity/source statement differs')
        # The wrapper freezes its numerical tree; its supervisor additionally
        # freezes the auditor, job list and canonical manifest. Some wrappers
        # use source_sha256 for incoming witness bytes, hence prefer _before.
        before = hashes(producer_summary.get('source_sha256_before', producer_summary.get('source_sha256')),
                        repo, 'Producer source')
        after = hashes(producer_summary['source_sha256_after'], repo, 'Producer end')
        require(before == after and required | {command_script} <= before.keys() and
                all(path in start and start[path] == digest for path, digest in before.items()),
                'Producer numerical start/end differs from supervisor')
        artifacts = component['component_artifacts_sha256']
        require(bool(artifacts), 'No component evidence bytes recorded')
        for name, expected in artifacts.items():
            path = (component_folder/name).resolve()
            require(path.is_relative_to(component_folder) and sha(path.read_bytes()) == expected,
                    'Component artifact changed or escaped its folder')
            supervised_artifacts[path] = expected
        checked_components.append({'role': component['role'], 'adopted': component['adopted'],
                                   'process_seconds': seconds, 'process_exit_code': job['returncode'],
                                   'folder': str(component_folder.relative_to(repo)).replace('\\', '/')})
        for path, digest in start.items():
            require(path not in all_sources or all_sources[path] == digest, 'Source identity conflict')
            all_sources[path] = digest
            require(path not in supervised_sources or supervised_sources[path] == digest, 'Supervised source identity conflict')
            supervised_sources[path] = digest
    require(len(models) == 1, 'Components have different physical models')
    model = next(iter(models))
    identities = {}
    for mode in ('lower', 'upper'):
        certificate = read(folder/f'best_{mode}_certificate.json')
        identities[mode] = _root(certificate, case, mode, policy, model)
        require(summary.get(mode) == certificate, 'Posthoc summary and adopted certificate differ')
        finals = [c for c in components if c['adopted'] and f'{mode}_adopted_' in c['role']]
        require(len(finals) == 1 and finals[0]['os_returncode'] == 0, 'Adopted closure did not finish successfully')
        original_folder = inside(finals[0]['folder'], repo)
        for suffix in ('mesh.npz', 'field.npz'):
            require((folder/f'best_{mode}_{suffix}').read_bytes() ==
                    (original_folder/f'best_{mode}_{suffix}').read_bytes(), 'Posthoc field/mesh differs from final solve')
        original = read(original_folder/f'best_{mode}_certificate.json')
        require(_root(original, case, mode, policy, model) == identities[mode] and
                original['audit']['discrete_root_search'] == certificate['audit']['discrete_root_search'],
                'Posthoc own root differs from adopted solver root')
        if audit is not None:
            side = audit['sides'][mode]
            require(all(side.get(k) == identities[mode][k] for k in ('fs', 'mesh_id', 'physical_model_id', 'root_complete')) and
                    side['audit']['discrete_root_search'] == certificate['audit']['discrete_root_search'],
                    'Independent pair audit differs from adopted original root')
            tracked = hashes(side.get('tracked_file_sha256', side.get('file_sha256')), repo, 'Independent audited bytes')
            required_evidence = {input_path, folder/f'best_{mode}_certificate.json',
                                 folder/f'best_{mode}_mesh.npz', folder/f'best_{mode}_field.npz'}
            require(required_evidence <= tracked.keys(), 'Independent audit omitted canonical input or adopted evidence')
            require(all(sha(path.read_bytes()) == tracked[path] for path in required_evidence),
                    'Independently audited evidence changed')
    timing = manifest['timing']
    adopted_seconds = sum(c['process_wall_seconds'] for c in components
                          if c['adopted'] and c['process_wall_seconds'] is not None)
    require(abs(adopted_seconds-timing['adopted_geometry_solve_component_process_sum_seconds']) < 1e-6 and
            abs(sum(c['process_wall_seconds'] for c in components if c['process_wall_seconds'] is not None)-
                timing['all_recipe_trial_process_sum_seconds']) < 1e-6,
            'Composite process-duration sum differs')
    if independent_dual_audits:
        unsupervised = [c for c in components if c.get('ownership_file') is None]
        require(len(unsupervised) == 1 and
                timing.get('unsupervised_repair_wrapper_seconds') == unsupervised[0]['wrapper_seconds'] and
                timing.get('unsupervised_repair_process_wall_seconds') is None,
                'Unsupervised repair wrapper/process timing must remain separate')
    for path, digest in all_sources.items():
        require(sha(path.read_bytes()) == digest, 'Source changed during composite verification')
    return {'source_provenance_verified': True, 'input_provenance_verified': True,
            'model_identity_verified': True, 'own_root_identity_verified': True,
            'evidence_kind': 'known_recipe_geometry_replay', 'historical_geometry_imported': True,
            'historical_field_imported': False, 'historical_fs_imported': False,
            'posthoc_composite': True, 'cold_execution_status_verified': False,
            'explicit_producer_end_source_map': True, 'components': checked_components,
            'verified_source_count': len(all_sources), 'physical_model_id': model, 'sides': identities,
            'component_manifest_sha256': sha(manifest_path.read_bytes()),
            'independent_original_dual_audits': independent_dual_audits,
            'summary_sha256': sha(summary_path.read_bytes()), 'timing': timing,
            'input_sha256': sha(input_path.read_bytes()), 'input_canonical_sha256': sha(canonical(record)),
            'source_closure_scope': ('Actual supervised component start/end maps plus one explicitly captured, '
                'unowned no-optimizer repair with actual producer maps and independent original-dual replay; '
                'no invented PID/process timing or combined start' if independent_dual_audits else
                'Actual per-component supervisor start/end maps; current complete numerical inventories; posthoc assembly has no invented combined start')}
