"""Verify cached producer provenance without importing solvers or running one.

This establishes code/input/case/root identity, not a new physical audit and
not successful worker termination. Execution status/budget remain separate.
"""
from __future__ import annotations
import ast
import hashlib
import json
import math
from pathlib import Path
import re

ROOT = Path(__file__).resolve().parent
REPO = ROOT.parents[1]
CATALOG = Path('docs/benchmarks/G1R17_native26_20261013/data/inputs')
HEX = re.compile(r'^[0-9a-f]{64}$')


def require(condition, message):
    if not condition:
        raise ValueError(message)


def sha(data):
    return hashlib.sha256(data).hexdigest()


def canonical(value):
    return json.dumps(value, sort_keys=True, separators=(',', ':'), ensure_ascii=False).encode('utf8')


def _pairs(values):
    result = {}
    for key, value in values:
        require(key not in result, 'Duplicate JSON key: ' + key)
        result[key] = value
    return result


def read(path):
    def invalid(value):
        raise ValueError('Nonfinite JSON constant: ' + value)
    return json.loads(path.read_text(encoding='utf8'), object_pairs_hook=_pairs, parse_constant=invalid)


def inside(name, repo):
    path = Path(name)
    path = (path if path.is_absolute() else repo/path).resolve()
    require(path.is_relative_to(repo), 'Provenance file is outside the repository: ' + str(name))
    return path


def hashes(value, repo, label):
    require(isinstance(value, dict) and bool(value), label + ': missing source hashes')
    result = {}
    for name, digest in value.items():
        require(isinstance(name, str) and isinstance(digest, str) and HEX.fullmatch(digest), label + ': invalid path/hash')
        path = inside(name, repo)
        require(path not in result, label + ': aliased source path')
        result[path] = digest
    return result


def finite_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool) and math.isfinite(value)


def _root(metadata, case, mode, policy, model_id):
    require(isinstance(metadata, dict), 'Missing ' + mode + ' certificate')
    root = metadata.get('audit', {}).get('discrete_root_search', {})
    fs = metadata.get('fs'); mesh = metadata.get('mesh_id')
    require(finite_number(fs) and fs > 0 and isinstance(mesh, str) and bool(HEX.fullmatch(mesh)), 'Invalid own-mesh endpoint')
    require(metadata.get('case') == case and metadata.get('mode') == mode and metadata.get('q') == 8 and
            metadata.get('physical_model_id') == model_id, 'Certificate case/mode/q/model differs from canonical producer input')
    identity = {'case':case, 'mode':mode, 'q':8, 'policy':policy, 'mesh_id':mesh, 'physical_model_id':model_id}
    require(all(root.get(key) == value for key, value in identity.items()), 'Root case/mode/q/policy/mesh/model differs')
    bracket = root.get('bracket')
    require(isinstance(bracket, list) and len(bracket) == 2 and all(finite_number(v) for v in bracket) and
            0 < bracket[0] < bracket[1], 'Invalid own-root bracket')
    require(bracket[1]-bracket[0] <= 1e-4*max(1., sum(bracket)/2), 'Original normalized root width fails')
    require(root.get('bracket_signs') == [1,-1] and metadata.get('root_complete') is True and root.get('root_complete') is True and
            root.get('endpoint') == fs == bracket[0 if mode == 'lower' else 1], 'Incomplete root signs/accepted endpoint')
    return {'fs':fs, 'mesh_id':mesh, 'physical_model_id':model_id, 'q':8, 'policy':policy,
            'root_complete':True, 'root_width_normalized':(bracket[1]-bracket[0])/max(1., sum(bracket)/2)}


def _model_id(record, builder):
    # This is the protocol implemented by the original mesh builder. E/nu,
    # water and boundary metadata remain additionally bound by full input SHA.
    base = sha(json.dumps(record['source_case'], sort_keys=True).encode())
    if not record['source_case'].get('pore'):
        return base
    rules = set()
    for node in ast.walk(ast.parse(builder.read_text(encoding='utf8'))):
        if isinstance(node, ast.Dict):
            for key, value in zip(node.keys, node.values):
                if isinstance(key, ast.Constant) and key.value == 'pressure_rule' and isinstance(value, ast.Constant) and isinstance(value.value, str):
                    rules.add(value.value)
    supported = {'declared_nodal_hydrostatic_P1_on_analysis_mesh', 'exact_declared_hydrostatic_waterline_conforming_v2'}
    require(len(rules) == 1 and rules <= supported, 'Unsupported or ambiguous coupled model identity rule')
    return sha(canonical({'base':base, 'water':record['water'], 'pressure_rule':next(iter(rules)), 'load_phase':'reference_total'}))


def verify_source_provenance(summary_path, case, *, repo=REPO, audit=None):
    """Return verified identities or raise ValueError/OSError on inconsistency.

    Sources are checked against their start manifest, producer summary maps,
    optional explicit checkpoint/end maps, and current complete code inventory.
    Older wrappers store only an unchanged-source statement at the end; report
    that limitation explicitly while independently checking all current bytes.
    """
    repo = Path(repo).resolve(); summary_path = inside(summary_path, repo)
    summary = read(summary_path); folder = summary_path.parent
    require(summary.get('case') == case, 'Summary case differs')
    require(summary.get('source_unchanged') is True, 'Producer did not report unchanged sources at checkpoint')
    start_path = folder/'source_manifest.json'
    start = hashes(read(start_path).get('source_sha256'), repo, 'Start manifest')
    environment = summary.get('environment', {})
    declared = hashes(environment.get('source_sha256'), repo, 'Summary source map')
    require(start == declared, 'Start and summary numerical source maps differ')
    for extra in (summary.get('wrapper_source_sha256'),):
        if extra is None:
            continue
        for path, digest in hashes(extra, repo, 'Wrapper source map').items():
            require(path not in declared or declared[path] == digest, 'Wrapper source identity conflicts')
            declared[path] = digest
    end_maps = []
    for holder in (environment, summary):
        for key in ('source_sha256_at_checkpoint', 'source_sha256_after'):
            if key in holder:
                value = hashes(holder[key], repo, 'Explicit end source map')
                require(value == start or value == declared, 'Start/end numerical source maps differ')
                end_maps.append(key)
    drivers = [path for path in declared if path.name == 'accuracy_driver.py' and path.parent.name == 'python']
    require(len(drivers) == 1, 'One producer implementation tree is required')
    python = drivers[0].parent; producer = python.parent
    required = {p.resolve() for p in python.rglob('*.py') if '__pycache__' not in p.parts}
    required |= {producer/'run_case.py', producer/'INPUTS_26.json'}
    require(required <= declared.keys(), 'Producer source map omits current numerical code/wrapper/input manifest: ' +
            ', '.join(sorted(str(p.relative_to(repo)) for p in required-declared.keys())))
    for path, expected in declared.items():
        require(path.is_file() and sha(path.read_bytes()) == expected, 'A producer source changed: ' + str(path.relative_to(repo)))
    manifest_path = producer/'INPUTS_26.json'; manifest = read(manifest_path)
    cases = manifest.get('cases', [])
    require(len(cases) == 26 and len({entry.get('case') for entry in cases}) == 26, 'Canonical manifest is not 26 unique cases')
    matches = [entry for entry in cases if entry.get('case') == case]
    require(len(matches) == 1, 'Case is absent/duplicated in canonical manifest')
    entry = matches[0]; input_path = inside(entry['input']['path'], repo)
    require(input_path == (repo/CATALOG/(case+'.json')).resolve(), 'Manifest case points at another canonical input')
    raw = input_path.read_bytes(); record = read(input_path)
    require(record.get('case') == case and record.get('source_case', {}).get('id') == case, 'Canonical record case differs')
    require(sha(raw.replace(b'\r\n', b'\n')) == entry['input']['sha256_lf'] and
            sha(canonical(record)) == entry['input']['sha256_canonical_json'], 'Canonical input differs from frozen producer manifest')
    policy = record['source_case']['policy']
    require(entry.get('policy') == policy, 'Canonical policy differs from producer manifest')
    if 'initial_mesh' in entry:
        initial = inside(entry['initial_mesh']['path'], repo)
        require(sha(initial.read_bytes()) == entry['initial_mesh']['sha256'], 'Declared initial catalog mesh changed')
    model = _model_id(record, python/'benchmark_models.py')
    identities = {}
    for mode in ('lower','upper'):
        metadata = read(folder/('best_'+mode+'_certificate.json'))
        identity = _root(metadata, case, mode, policy, model)
        producer_side = summary.get(mode)
        require(isinstance(producer_side, dict), 'Summary missing an adopted side')
        for key in ('case','mode','q','fs','mesh_id','physical_model_id','root_complete'):
            require(producer_side.get(key) == metadata.get(key), 'Summary and adopted certificate differ: '+mode+'/'+key)
        if audit is not None:
            require(audit.get('case') == case, 'Independent audit case differs')
            side = audit.get('sides', {}).get(mode, {})
            for key in ('fs','mesh_id','physical_model_id','root_complete'):
                require(side.get(key) == identity.get(key), 'Audit and producer certificate identity differ: '+mode+'/'+key)
            audit_root = side.get('audit', side).get('discrete_root_search')
            require(audit_root == metadata['audit']['discrete_root_search'], 'Audit and producer root records differ')
            tracked = side.get('tracked_file_sha256', side.get('file_sha256', {}))
            paths = {inside(name, repo):digest for name,digest in tracked.items()}
            required_evidence = {input_path, folder/('best_'+mode+'_certificate.json')}
            for key, fallback in (('mesh_path','best_'+mode+'_mesh.npz'),('field_path','best_'+mode+'_field.npz')):
                evidence = (folder/metadata.get(key,fallback)).resolve()
                require(evidence.is_relative_to(folder), 'Producer evidence points outside its own case folder')
                required_evidence.add(evidence)
            require(required_evidence <= paths.keys(), 'Independent audit omitted canonical input or adopted evidence')
            for path in required_evidence:
                require(sha(path.read_bytes()) == paths[path], 'An independently audited adopted file changed')
        identities[mode] = identity
    # Repeat the whole source comparison after input/certificate processing.
    for path, expected in declared.items():
        require(sha(path.read_bytes()) == expected, 'Producer source changed during provenance verification')
    normalized = {str(p.relative_to(repo)).replace('\\','/'):h for p,h in sorted(declared.items())}
    return {'source_provenance_verified':True, 'input_provenance_verified':True, 'model_identity_verified':True,
            'own_root_identity_verified':True, 'verified_source_count':len(declared),
            'verified_source_set_sha256':sha(canonical(normalized)), 'producer_root':str(producer.relative_to(repo)).replace('\\','/'),
            'producer_source_manifest_sha256':sha(start_path.read_bytes()), 'summary_sha256':sha(summary_path.read_bytes()),
            'input_sha256':sha(raw), 'input_canonical_sha256':sha(canonical(record)), 'input_manifest_sha256':sha(manifest_path.read_bytes()),
            'physical_model_id':model, 'sides':identities, 'explicit_producer_end_source_map':bool(end_maps),
            'source_closure_scope':'start manifest + summary maps + current complete file set; explicit end map when available',
            'cold_execution_status_verified':False,
            'execution_note':'Worker exit status, supervisor termination and budget are separate from scientific/source provenance.'}
