"""Canonical 26-case benchmark models, including the declared D08A water field.

No field or root is imported by load_model. D08A uses nodal hydrostatic
pressure with an exactly represented kink at the waterline. Every admissible
mesh therefore represents the same prescribed continuum body force. Refinement must reconstruct that
load, rather than silently return the dry body force.
"""
from __future__ import annotations

import argparse
import copy
import hashlib
import json
from pathlib import Path
import struct

import experiment_core as core
from experiment_core import np

REPO = Path(__file__).resolve().parents[3]
RUN_ROOT = Path(__file__).resolve().parents[1]
MANIFEST_PATH = RUN_ROOT / 'INPUTS_26.json'
_BASE_BUILDER = getattr(core, '_benchmark26_base_builder', core.mesh_from_arrays)
_MANIFEST = None


def _sha(raw):
    return hashlib.sha256(raw).hexdigest()


def _canonical(record):
    return json.dumps(record, sort_keys=True, separators=(',', ':'), ensure_ascii=False).encode('utf8')


def manifest():
    global _MANIFEST
    if _MANIFEST is None:
        data = json.loads(MANIFEST_PATH.read_text(encoding='utf8'))
        ids = [entry['case'] for entry in data['cases']]
        if len(ids) != 26 or len(set(ids)) != 26:
            raise ValueError('The canonical manifest must contain exactly 26 unique cases')
        _MANIFEST = data
    return _MANIFEST


def case_entry(case):
    for entry in manifest()['cases']:
        if entry['case'] == case:
            return entry
    raise ValueError(f'Unknown benchmark case: {case}')


def _record(case):
    entry = case_entry(case)
    raw = (REPO / entry['input']['path']).read_bytes()
    # Windows checkout may convert LF to CRLF. Both exact byte hashes are
    # recorded; only that newline conversion is allowed, never changed values.
    if _sha(raw.replace(b'\r\n', b'\n')) != entry['input']['sha256_lf']:
        raise ValueError(f'{case}: canonical input bytes differ from the manifest')
    record = json.loads(raw.decode('utf8'))
    if _sha(_canonical(record)) != entry['input']['sha256_canonical_json']:
        raise ValueError(f'{case}: canonical input values differ from the manifest')
    _validate_record(record)
    return record


def _validate_record(record):
    case = record.get('case')
    entry = case_entry(case)
    if _sha(_canonical(record)) != entry['input']['sha256_canonical_json']:
        raise ValueError(f'{case}: refinement record differs from the frozen model')
    if record['source_case']['policy'] != entry['policy']:
        raise ValueError(f'{case}: analysis policy differs from the manifest')
    coupling = bool(record.get('water', {}).get('coupling'))
    if bool(record['source_case'].get('pore')) != coupling:
        raise ValueError(f'{case}: pore-pressure flags disagree')
    if coupling:
        water = record['water']
        if (case != 'D08A' or water.get('mode') != 2 or water.get('gamma_w') != 9.81 or
            water.get('suction') != 0 or water.get('water_table') != [[0., 3.], [20., 3.]] or
            water.get('catalog_spec', {}).get('formula') != '9.81*max(3-y,0)'):
            raise ValueError('BLOCKED: unsupported coupled water model; no dry fallback is allowed')
        # There is no declared external hydrostatic pressure on an exposed
        # ground face in this particular case. The other wet faces are fixed.
        if water['catalog_spec'].get('surface_water_traction') != 'none; exposed ground is above y=3':
            raise ValueError('BLOCKED: D08A external-water traction definition differs')


def hydrostatic_body(mesh, record):
    """Exact common prescribed hydrostatic field on a waterline-conforming mesh."""
    from waterline_conform import exact_hydrostatic_body
    if not record.get('water', {}).get('coupling'):
        body=np.zeros_like(mesh.body)
        body[:,1,1]=-mesh.materials[mesh.material_id,3]
        return body
    return exact_hydrostatic_body(mesh,level=3.,gamma_water=record['water']['gamma_w'])


def mesh_from_arrays(xy, tri, ids, record):
    _validate_record(record)
    # The legacy builder uses an absolute 1e-9 area assertion, which rejects
    # D05A's L*1000 similitude for a 5.96e-8 floating summation error (relative
    # 3.8e-16). Check the canonical areas independently before supplying the
    # identical computed sums to that internal assertion. No coordinate,
    # material, load or finite-element operator is changed.
    points = np.asarray(xy)[np.asarray(tri)]
    a = points[:, 1] - points[:, 0]
    b = points[:, 2] - points[:, 0]
    products = np.stack((a[:, 0] * b[:, 1], a[:, 1] * b[:, 0]), axis=1)
    areas = (products[:, 0] - products[:, 1]) / 2
    actual = np.bincount(np.asarray(ids, dtype=int), weights=areas,
                         minlength=len(record['region_areas']))
    expected = np.asarray(record['region_areas'], dtype=float)
    if actual.shape != expected.shape or not np.isfinite(areas).all() or np.any(areas <= 0):
        raise ValueError('Invalid triangle areas or material region indices')
    operations = 8 * max(1, len(areas)) * np.finfo(float).eps
    if operations >= 1:
        raise ValueError('Geometry is too large for the area roundoff guard')
    summation_bound = operations / (1 - operations) * max(1., float(np.sum(abs(products)) / 2))
    area_limit = max(1e-9, summation_bound)
    area_error = float(np.max(abs(actual - expected), initial=0.))
    if area_error > area_limit:
        raise ValueError(f'Canonical material area differs: {area_error} > {area_limit}')
    construction_record = dict(record, region_areas=actual.tolist())
    mesh = _BASE_BUILDER(xy, tri, ids, construction_record)
    mesh.provenance['canonical_region_area_error'] = area_error
    mesh.provenance['canonical_region_area_roundoff_limit'] = area_limit
    if record.get('water', {}).get('coupling'):
        mesh.body = hydrostatic_body(mesh, record)
        identity = {'base': mesh.physical_model_id, 'water': record['water'],
                    'pressure_rule': 'exact_declared_hydrostatic_waterline_conforming_v2',
                    'load_phase': 'reference_total'}
        mesh.physical_model_id = _sha(_canonical(identity))
        mesh.provenance['water_model'] = identity['pressure_rule']
        mesh.provenance['water_level_m'] = 3.
        mesh.provenance['gamma_water'] = 9.81
    entry = case_entry(record['case'])
    mesh.provenance['benchmark_input_sha256_lf'] = entry['input']['sha256_lf']
    mesh.provenance['benchmark_input_canonical_sha256'] = entry['input']['sha256_canonical_json']
    return mesh.validate()


def install_model_builder():
    """Cover cloned-core refinement and saved-mesh loading with one builder."""
    if core.mesh_from_arrays is mesh_from_arrays:
        return
    current_base = getattr(core, '_benchmark26_base_builder', None)
    if current_base is not None and current_base is not _BASE_BUILDER:
        raise ValueError('A different benchmark model builder is already installed')
    core._benchmark26_base_builder = _BASE_BUILDER
    core.mesh_from_arrays = mesh_from_arrays


def load_model(case):
    record = _record(case)
    entry = case_entry(case)
    path = REPO / entry['initial_mesh']['path']
    raw = path.read_bytes()
    if _sha(raw) != entry['initial_mesh']['sha256']:
        raise ValueError(f'{case}: initial mesh differs from the manifest')
    if len(raw) < 8:
        raise ValueError('Truncated geometry header')
    nn, ne = struct.unpack_from('<ii', raw)
    if nn <= 0 or ne <= 0 or len(raw) != 8 + 16 * nn + 16 * ne:
        raise ValueError('Invalid initial mesh dimensions')
    xy = np.frombuffer(raw, '<f8', 2 * nn, 8).copy().reshape(nn, 2)
    cells = np.frombuffer(raw, '<i4', 4 * ne, 8 + 16 * nn).copy().reshape(ne, 4)
    mesh = mesh_from_arrays(xy, cells[:, :3], cells[:, 3], record)
    mesh.provenance['initial_mesh_source'] = entry['initial_mesh']['path']
    mesh.provenance['initial_mesh_alias_case'] = entry['initial_mesh'].get('alias_case')
    return mesh, copy.deepcopy(record)


def _read_native_state(path):
    raw = path.read_bytes()
    if raw[:8] != b'RP25M001':
        raise ValueError('Unknown native26 witness format')
    position = 8
    def take(dtype, count):
        nonlocal position
        size = np.dtype(dtype).itemsize * count
        if count < 0 or position + size > len(raw):
            raise ValueError('Truncated native26 witness')
        result = np.frombuffer(raw, dtype, count, position).copy()
        position += size
        return result
    nn, ne, nm, nr, q = map(int, take('<i4', 5))
    if min(nn, ne, nm, q) <= 0 or nr != 0:
        raise ValueError('Only finite soil-only native26 snapshots are supported')
    scale, factor, fs, objective, native_audit = map(float, take('<f8', 5))
    xy = take('<f8', 2 * nn).reshape(nn, 2)
    cells = take('<i4', 5 * ne).reshape(ne, 5)
    materials = take('<f8', 4 * nm).reshape(nm, 4)
    body = take('<f8', 4 * ne).reshape(ne, 2, 2)
    boundary = {}
    for _ in range(int(take('<i4', 1)[0])):
        a, b, kind, rigid = map(int, take('<i4', 4))
        key = tuple(sorted((a, b)))
        if not 0 <= kind < 4 or key in boundary:
            raise ValueError('Invalid native26 boundary')
        boundary[key] = {'kind': ['load', 'fixed', 'roller_x', 'roller_y'][kind],
                         'rigid': rigid, 'loads': take('<f8', 4).reshape(2, 2).T}
    field = take('<f8', int(take('<i4', 1)[0]))
    if position != len(raw) or factor != fs or np.any(cells[:, 4] != -1):
        raise ValueError('Native26 phase/rigid/trailing-data mismatch')
    if not np.isfinite([scale, factor, fs, objective, native_audit]).all() or fs <= 0:
        raise ValueError('Nonfinite native26 metadata')
    return {'xy': xy, 'cells': cells, 'materials': materials, 'body': body,
            'boundary': boundary, 'field': field, 'q': q, 'fs': fs,
            'objective': objective, 'scale': scale, 'sha256': _sha(raw)}


def verify_legacy_nonconforming_models_not_applicable(native_root, audit_fields=True):
    """No numerical solve: verify 26 geometries and both exact coupled states."""
    install_model_builder()
    cases = []
    for entry in manifest()['cases']:
        mesh, record = load_model(entry['case'])
        area, _, _ = core.topology(mesh)
        regions = np.bincount(mesh.material_id, weights=area, minlength=len(mesh.materials))
        cases.append({'case': entry['case'], 'nodes': len(mesh.xy), 'elements': len(mesh.tri),
                      'policy': mesh.provenance['psi_policy'],
                      'physical_model_id': mesh.physical_model_id,
                      'region_area_error': float(np.max(abs(regions - record['region_areas']))),
                      'pore': bool(record['water']['coupling'])})
    record = _record('D08A')
    coupled = []
    for side in ['lower', 'upper']:
        path = native_root / 'cases' / 'D08A' / f'witness_{side}.bin'
        state = _read_native_state(path)
        mesh = mesh_from_arrays(state['xy'], state['cells'][:, :3], state['cells'][:, 3], record)
        if not np.array_equal(mesh.materials, state['materials']) or mesh.stress_scale != state['scale']:
            raise ValueError('D08A native materials or scale differ')
        if mesh.boundary.keys() != state['boundary'].keys():
            raise ValueError('D08A native boundary keys differ')
        for key, value in mesh.boundary.items():
            old = state['boundary'][key]
            if value['kind'] != old['kind'] or value['rigid'] != old['rigid'] or not np.array_equal(value['loads'], old['loads']):
                raise ValueError('D08A native boundary load or restraint differs')
        error = float(np.max(abs(mesh.body - state['body'])))
        if error > 1e-9 or state['q'] != 8:
            raise ValueError('D08A native body force or quadrature differs')
        check = {'side': side, 'source': str(path), 'sha256': state['sha256'],
                 'elements': len(mesh.tri), 'fs': state['fs'], 'q': state['q'],
                 'body_max_error': error, 'materials_equal': True, 'boundary_equal': True,
                 'solver_calls': 0, 'root_imported': False}
        if audit_fields:
            result = core.evaluate(core.Prepared(mesh, state['fs'], side, 8, reduction='none'), state['field'])
            check['original_q8_audit'] = result
            if not result['physical_pass'] or not result['witness']:
                raise ValueError('D08A saved field fails the current original q8 gate')
        # Verify that core refinement and subsequent reconstructions keep the
        # water rule. No saved field or root is transferred by this check.
        child = core.uniform_refine(mesh, record)
        expected = hydrostatic_body(child, record)
        if not np.array_equal(child.body, expected) or child.physical_model_id != mesh.physical_model_id:
            raise ValueError('D08A refinement lost its coupled pressure model')
        check['refined_elements'] = len(child.tri)
        check['refined_body_rebuilt_exactly'] = True
        coupled.append(check)
    return {'status': 'PASS', 'scope': 'geometry/model reconstruction and original saved-field audits; no numerical solve',
            'geometry_cases': cases, 'D08A': coupled,
            'manifest_sha256': _sha(MANIFEST_PATH.read_bytes()),
            'model_builder_sha256': _sha(Path(__file__).read_bytes())}


if __name__ == '__main__':
    raise SystemExit('Use verify_water_model.py: legacy nonconforming saved fields are not transferable to this prototype')
