"""Independent, solve-free audit of final cold Excel own-mesh checkpoints.

Fs, fields, and exported layouts are used without normalization or backoff.
RP25M001 and AC18P001 are decoded here rather than trusting native summaries.
Root flags/width are checked separately from independently rebuilt primal gates;
missing original endpoint dual vectors are explicitly reported, never invented.
"""
from __future__ import annotations
import argparse, collections, csv, hashlib, json, math, struct, sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
REPO = HERE.parents[1]
PYTHON = REPO / 'validation/G1R19_python26_20261016/python'
sys.path.insert(0, str(PYTHON))
import experiment_core as core
from experiment_core import np, sparse
from model_support.schema import topology


def sha(raw):
    return hashlib.sha256(raw).hexdigest()


def require(value, reason):
    if not value:
        raise ValueError(reason)


def exact(value):
    require(len(value) == 16, 'Invalid IEEE754_LE_HEX length')
    result = struct.unpack('<d', bytes.fromhex(value))[0]
    require(math.isfinite(result), 'Nonfinite exact double')
    return result


def decoded(raw):
    return raw.decode('utf16' if raw[:2] in (b'\xff\xfe', b'\xfe\xff') else 'utf8')


class Reader:
    def __init__(self, raw, magic):
        require(raw[:8] == magic, 'Unsupported binary magic')
        self.raw, self.pos = raw, 8

    def take(self, dtype, count):
        require(count >= 0 and self.pos + np.dtype(dtype).itemsize * count <= len(self.raw), 'Truncated binary')
        result = np.frombuffer(self.raw, dtype=dtype, count=count, offset=self.pos).copy()
        self.pos += result.nbytes
        return result

    def end(self):
        require(self.pos == len(self.raw), 'Unexpected trailing binary bytes')


def state(raw):
    r = Reader(raw, b'RP25M001')
    nn, ne, nm, nr, q = map(int, r.take('<i4', 5))
    require(min(nn, ne, nm) > 0 and nr == 0 and q == 8, 'Soil-only q8 required')
    scale, factor, fs, objective, native_audit = map(float, r.take('<f8', 5))
    require(all(map(math.isfinite, (scale, factor, fs, objective, native_audit))) and scale > 0 and factor == fs and fs > 0,
            'Invalid physical checkpoint metadata')
    xy = r.take('<f8', 2 * nn).reshape(nn, 2)
    cells = r.take('<i4', 5 * ne).reshape(ne, 5)
    materials = r.take('<f8', 4 * nm).reshape(nm, 4)
    body = r.take('<f8', 4 * ne).reshape(ne, 2, 2)
    boundary = {}
    for _ in range(int(r.take('<i4', 1)[0])):
        a, b, kind, rigid = map(int, r.take('<i4', 4))
        require(0 <= kind <= 3 and 0 <= min(a, b) < max(a, b) < nn and rigid == -1, 'Invalid exterior face')
        key = tuple(sorted((a, b)))
        require(key not in boundary, 'Duplicate exterior face')
        boundary[key] = dict(kind=['load', 'fixed', 'roller_x', 'roller_y'][kind], rigid=rigid,
                             loads=r.take('<f8', 4).reshape(2, 2).T)
    field = r.take('<f8', int(r.take('<i4', 1)[0]))
    r.end()
    require(np.isfinite(field).all() and np.all(cells[:, 4] == -1), 'Nonfinite field or unsupported rigid cell')
    return dict(xy=xy, cells=cells, materials=materials, body=body, boundary=boundary, field=field,
                q=q, scale=scale, fs=fs, objective=objective, native_audit=native_audit)


def metadata(path):
    result = collections.defaultdict(list)
    for line in path.read_text(encoding='utf8').splitlines():
        if line.startswith('original_model='):
            result['original_model'].append(line.split('=', 1)[1])
            continue
        for item in line.split(';'):
            key, value = item.split('=', 1)
            result[key].append(value)
    for key, values in result.items():
        require(len(values) == 1 or key in ('lower_stress_start', 'reaction_part'), 'Duplicate metadata key: ' + key)
    return result


def one(meta, key):
    require(key in meta and len(meta[key]) == 1, 'Missing/duplicate metadata: ' + key)
    return meta[key][0]


def original_model_matches(model, s, mesh):
    """Check the exact native model string against binary and rebuilt topology."""
    parts = model.split('|')
    nn, ne, nm = len(s['xy']), len(s['cells']), len(s['materials'])
    area, _, edges = topology(mesh)
    header = parts[0].split(',')
    require(header[:7] == list(map(str, (nn, ne, nm, 0, len(edges)))) + ['fixed', 'reference_total'], 'Native model header differs')
    require(exact(header[7]) == s['scale'] and header[8:] == ['ASSOCIATED', 'raw_audit=True'], 'Native model policy/scale/audit differs')
    cursor = 1
    for point in s['xy']:
        require(np.array_equal([exact(parts[cursor][:16]), exact(parts[cursor][16:])], point), 'Native model node differs')
        cursor += 1
    for cell, body in zip(s['cells'], s['body']):
        fields = parts[cursor].split(',')
        require(np.array_equal(list(map(int, fields[:5])), cell), 'Native model cell differs')
        packed = np.array([[exact(v[:16]), exact(v[16:])] for v in fields[5:]])
        require(np.array_equal(packed, body), 'Native model body differs')
        cursor += 1
    for material in s['materials']:
        require(np.array_equal([exact(parts[cursor][i:i+16]) for i in range(0, 64, 16)], material), 'Native model material differs')
        cursor += 1
    expected = {(edge.a, edge.b): edge for edge in edges}
    seen = set()
    for text in parts[cursor:cursor + len(edges)]:
        f = text.split(','); a, b, e, other = map(int, f[:4]); key = tuple(sorted((a, b)))
        require(key in expected and key not in seen, 'Native model edge differs')
        edge = expected[key]; seen.add(key)
        require(sorted([e, other]) == sorted(edge.elements) and f[4] == edge.kind and int(f[5]) == edge.rigid, 'Native model topology/restraint differs')
        loads = np.array([[exact(v[:16]), exact(v[16:])] for v in f[6:]])
        require(np.array_equal(loads, edge.loads), 'Native model edge load differs')
    require(len(seen) == len(edges), 'Native model edge omitted')
    constraints = [text for text in parts[cursor + len(edges):] if text]
    constraint_edges = set()
    for text in constraints:
        key, a, b = text.split(','); a, b = int(a), int(b)
        require(key == ':'.join(map(str, sorted((a, b)))) and tuple(sorted((a, b))) in mesh.boundary,
                'Unexpected native mesh constraint')
        require(key not in constraint_edges, 'Duplicate native mesh constraint')
        constraint_edges.add(key)
    require(len(constraint_edges) == len(mesh.boundary), 'Native boundary mesh constraint omitted')
    return dict(exact_binary_model_and_rebuilt_topology_match=True, edges=len(edges),
                exact_native_model_sha256=sha(model.encode('utf8')),
                boundary_mesh_constraints=len(constraint_edges))


def root_contract(meta, side, fs):
    require(one(meta, side + '_root_valid') == 'True' and one(meta, 'search_incomplete') == 'False', 'Incomplete native root')
    lo, hi = [exact(one(meta, side + '_root_' + end + '_exact')) for end in ('a', 'b')]
    require(0 < lo < hi and fs == (lo if side == 'lower' else hi), 'Checkpoint is not its own accepted root endpoint')
    width = (hi - lo) / max(1., (hi + lo) / 2)
    require(width <= 1e-4, 'Original normalized root width contract fails')
    return dict(a=lo, b=hi, width_absolute=hi-lo, width_normalized=width,
                original_formula='(b-a)/max(1,(a+b)/2) <= 1e-4', metadata_complete=True,
                endpoint_sign_scope='Native root flags and own-model endpoint identity; independent primal at accepted endpoint only')


def checkpoint(folder, name, side, rec):
    binpath = folder / 'checkpoints' / (name + '.bin')
    metapath = binpath.with_suffix('.meta.tsv')
    s, meta = state(binpath.read_bytes()), metadata(metapath)
    require(one(meta, 'schema') == 'OWN18C01' and one(meta, 'side') == side, 'Checkpoint side/schema differs')
    require(one(meta, 'policy') == 'ASSOCIATED' and one(meta, 'phase') == 'reference_total' and int(one(meta, 'epoch')) == 1, 'Checkpoint policy/phase/epoch differs')
    require(int(one(meta, 'q')) == s['q'] and int(one(meta, 'nodes')) == len(s['xy']) and int(one(meta, 'elements')) == len(s['cells']), 'Checkpoint size/q differs')
    require(exact(one(meta, 'value_exact')) == s['fs'] and exact(one(meta, 'objective_exact')) == s['objective'], 'Checkpoint exact Fs/objective differs')
    mesh = core.mesh_from_arrays(s['xy'], s['cells'][:, :3], s['cells'][:, 3], rec)
    require(np.array_equal(s['materials'], mesh.materials) and np.array_equal(s['body'], mesh.body), 'Canonical material/body load differs')
    require(s['scale'] == mesh.stress_scale and s['boundary'].keys() == mesh.boundary.keys(), 'Canonical scale/boundary set differs')
    for edge in mesh.boundary:
        require(all(s['boundary'][edge][k] == mesh.boundary[edge][k] for k in ('kind', 'rigid')) and
                np.array_equal(s['boundary'][edge]['loads'], mesh.boundary[edge]['loads']), 'Canonical boundary condition/load differs')
    model_check = original_model_matches(one(meta, 'original_model'), s, mesh)
    ne = len(mesh.tri)
    if side == 'lower':
        starts = dict(map(int, item.split(',')) for item in meta['lower_stress_start'])
        load_id = int(one(meta, 'lower_load_variable'))
        require(set(starts) == set(range(ne)) and len(s['field']) == 18*ne+1 and load_id == 18*ne and
                all(starts[e] == 18*e for e in starts), 'Unsupported lower variable layout')
        field = np.r_[np.concatenate([s['field'][starts[e]:starts[e]+18] for e in range(ne)]), s['field'][load_id]]
    else:
        require(len(s['field']) == 12*ne, 'Unsupported upper physical velocity layout')
        field = s['field'].copy()
    before = sha(field.tobytes())
    pr = core.Prepared(mesh, s['fs'], side, 8)
    audit = core.evaluate(pr, field)
    require(before == sha(field.tobytes()) and np.array_equal(field, s['field']), 'Audit mutated/scaled source field')
    require(audit['physical_pass'] and audit['witness'], 'Original q8 independent primal gate failed')
    # Native lower state.objective stores the maximized load multiplier;
    # Python minimizes -lambda. Upper stores the dissipation objective.
    native_comparable_objective = -audit['objective'] if side == 'lower' else audit['objective']
    require(abs(native_comparable_objective - s['objective']) <= 1e-12, 'Original objective differs')
    area, _, _ = topology(mesh)
    force = (area[:, None] * mesh.body[:, :, 1]).sum(axis=0)
    result = dict(side=side, checkpoint=binpath.relative_to(REPO).as_posix(), checkpoint_sha256=sha(binpath.read_bytes()),
                  metadata_sha256=sha(metapath.read_bytes()), fs=s['fs'], q=8, nodes=len(mesh.xy), elements=ne,
                  source=one(meta, 'source'), mesh_id=mesh.mesh_id, physical_model_id=mesh.physical_model_id,
                  policy_native='ASSOCIATED', policy_python=rec['source_case']['policy'], phase=mesh.load_phase,
                  original_model_check=model_check, canonical_material_body_and_boundary_exact=True,
                  area=float(sum(area)), reference_body_force=force.tolist(), field_sha256=before,
                  source_field_unchanged=True, native_objective=s['objective'],
                  objective_convention='Native lower maximizes lambda; Python minimizes -lambda' if side == 'lower' else 'Both minimize dissipation minus fixed-load work',
                  independent_original_q8=audit, root=root_contract(meta, side, s['fs']))
    return result, one(meta, 'original_model')


def replay(path):
    r = Reader(path.read_bytes(), b'AC18P001')
    n, m, nnz, nc, neq = map(int, r.take('<i4', 5))
    require(min(n, m, nnz, nc) > 0 and 0 <= neq <= m, 'Invalid original CSR dimensions')
    ptr, ids, val = r.take('<i4', m+1), r.take('<i4', nnz), r.take('<f8', nnz)
    rhs, cost = r.take('<f8', m), r.take('<f8', n)
    cones = r.take('<i4', 2*nc).reshape(nc, 2)
    fs = float(r.take('<f8', 1)[0]); q = int(r.take('<i4', 1)[0])
    dual = '.accuracy_dual_' in path.name
    field = r.take('<f8', m if dual else n); r.end()
    require(q == 8 and all(np.isfinite(a).all() for a in (val, rhs, cost, field)) and np.all(np.diff(ptr) >= 0) and
            ptr[0] == 0 and ptr[-1] == nnz and np.all((ids >= 0) & (ids < n)), 'Invalid original CSR/q/finite data')
    A = sparse.csr_matrix((val, ids, ptr), shape=(m, n))
    slack = field if dual else rhs-A@field
    stat = float(np.max(abs(cost+A.T@field), initial=0.)) if dual else float(np.max(abs(slack[:neq]), initial=0.))
    require(np.array_equal(cones[0], [0, neq]), 'Invalid equality cone')
    offset, cone = neq, 0.
    for kind, dim in cones[1:]:
        block = slack[offset:offset+dim]
        # Native cone.type uses its dimension (3) for a Lorentz cone.
        require((kind == 1 and dim == 1) or (kind == 3 and dim == 3), 'Unsupported original cone')
        cone = max(cone, float(-block[0] if kind == 1 else np.linalg.norm(block[1:])-block[0]))
        offset += dim
    require(offset == m and stat <= 1e-7 and cone <= (1e-8 if dual else 1e-7), 'Literal original native proof gate failed')
    context = {}
    context_raw = path.with_suffix(path.suffix+'.context.tsv').read_bytes()
    context_text = decoded(context_raw)
    for line in context_text.splitlines():
        split = line.split('\t')
        if len(split) == 2 and split[0] in ('model', 'fs', 'q'): context[split[0]] = split[1]
    require(exact(context['fs']) == fs and int(context['q']) == q, 'Original proof context Fs/q differs')
    native_cells = int(context['model'].split(',', 2)[1])
    is_lower = path.name.endswith('_lower.bin')
    if is_lower:
        expected_cost = np.zeros(18*native_cells+1); expected_cost[-1] = -1.
        require(n == len(expected_cost) and np.array_equal(cost, expected_cost), 'Original free-load lower objective/layout required')
    else:
        require(n == 12*native_cells and not np.any(cost), 'D02 original zero-cost upper objective/layout required')
    return dict(path=path.relative_to(REPO).as_posix(), sha256=sha(path.read_bytes()), context_sha256=sha(path.with_suffix(path.suffix+'.context.tsv').read_bytes()),
                mode='lower' if is_lower else 'upper', dual=dual, q=q, fs=fs, variables=n, rows=m,
                raw_equality_or_stationarity=stat, raw_cone=cone,
                original_objective=float(-rhs@field if dual else cost@field),
                exact_native_model_sha256=sha(context['model'].encode('utf8')), all_original_gates_pass=True)


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--run', type=Path, default=REPO/'versions/G1R19_2pct/results/owned_excel_1792149830248919100')
    parser.add_argument('--case', default='D02A')
    parser.add_argument('--lower', default='00004_lower')
    parser.add_argument('--upper', default='00005_upper')
    parser.add_argument('--output', type=Path, default=HERE/'COLD_NATIVE_D02_INDEPENDENT_AUDIT.json')
    args = parser.parse_args(); folder = args.run.resolve()
    record_path = REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs'/f'{args.case}.json'
    record = json.loads(record_path.read_text(encoding='utf8'))
    require(not record['source_case']['pore'] and record['source_case']['policy'] == 'equal_reduced_phi', 'This independent checkpoint auditor is limited to dry associated cases')
    require(json.loads((folder/'input_record.json').read_text(encoding='utf8')) == record, 'Cold actual input differs from canonical input')
    summary = json.loads((folder/'summary.json').read_text(encoding='utf8'))
    require(summary['status'] == 'PASS' and summary['owned_excel_exited'] is True, 'Native run did not finish/exit')
    require(summary['case'] == args.case and summary['imported_field_or_root'] is False and summary['solver_threads'] == 1,
            'Cold case/input/thread provenance differs')
    require(summary['input_sha256'] == sha(record_path.read_bytes()) and summary['dll_sha256'] == sha((folder/'rpfem_native_x64.dll').read_bytes()),
            'Canonical input/DLL bytes differ from frozen native run')
    for name, expected in summary['source_sha256'].items():
        require(sha((folder/'source_snapshot'/(name+'.bas')).read_bytes()) == expected, 'Frozen native source snapshot differs: '+name)
    results = {}
    for side in ('lower', 'upper'):
        results[side], _ = checkpoint(folder, getattr(args, side), side, record)
    lower, upper = results['lower']['fs'], results['upper']['fs']
    cold = summary['cold_result']
    require(cold[:4] == ['AUDITED', True, lower, upper] and cold[4:10] == [results['lower']['elements'], results['upper']['elements'], 8, 8, True, True] and
            cold[10:14] == [results['lower']['root']['a'], results['lower']['root']['b'], results['upper']['root']['a'], results['upper']['root']['b']] and cold[14] is False,
            'Final native cold result differs from adopted own-mesh checkpoints')
    gap = 2*(upper-lower)/(upper+lower)
    require(0 < lower <= upper and gap <= .02, 'Certified bound gap exceeds 2%')
    proofs = [replay(path) for path in sorted((folder/'RPFEM_logs').glob('*.accuracy_*.bin'))]
    require(proofs, 'Original proof exports missing')
    for side, result in results.items():
        identity = result['original_model_check']['exact_native_model_sha256']
        result['matching_literal_endpoint_proofs'] = [p['path'] for p in proofs if p['mode'] == side and p['fs'] == result['fs'] and p['exact_native_model_sha256'] == identity]
        result['root']['literal_proofs_by_endpoint'] = {key:[p['path'] for p in proofs if p['mode'] == side and p['fs'] == result['root'][key] and p['exact_native_model_sha256'] == identity] for key in ('a', 'b')}
    logs = list((folder/'RPFEM_logs').glob('*analysis.tsv'))
    require(len(logs) == 1, 'Ambiguous production analysis log')
    rows = list(csv.DictReader(decoded(logs[0].read_bytes()).splitlines(), delimiter='\t'))
    ends = [r for r in rows if r['event'] == 'solve_end']
    log_summary = dict(path=logs[0].relative_to(REPO).as_posix(), sha256=sha(logs[0].read_bytes()),
                       solve_start_count=sum(r['event'] == 'solve_start' for r in rows),
                       solve_end_count=len(ends),
                       calls_by_phase_mode=[dict(phase=k[0],mode=k[1],calls=v) for k,v in collections.Counter((r['candidate'],r['mode']) for r in ends).items()],
                       run_end_present=any(r['event'] == 'run_end' for r in rows),
                       upper_wedge_endpoints=[r for r in ends if r['candidate'] == 'upper_c0_wedge'])
    require(log_summary['run_end_present'] and log_summary['solve_start_count'] == len(ends), 'Production trace is incomplete')
    output = dict(status='INDEPENDENT_PRIMAL_Q8_AND_NATIVE_ROOT_CONTRACT_PASS', case=args.case,
                  scope='Fresh Excel adopted own-mesh fields independently rebuilt at exact saved Fs; no solver calls, no imported historical fields, no field normalization or Fs backoff',
                  gap=gap, gap_percent=100*gap, target_gap=.02, lower=lower, upper=upper,
                  sides=results, literal_original_proofs=proofs, literal_original_proof_count=len(proofs), production_log=log_summary,
                  root_proof_limit='Root metadata/model/width and accepted physical endpoint are independently checked. Complete original dual/ray vectors at every rejected root endpoint are not exported, so the native opposite signs are not independently re-proved here.',
                  input_sha256=sha(record_path.read_bytes()), input_canonical_sha256=sha(json.dumps(record,sort_keys=True,separators=(',',':'),ensure_ascii=False).encode('utf8')),
                  run=folder.relative_to(REPO).as_posix(), native_summary_sha256=sha((folder/'summary.json').read_bytes()),
                  native_cold_result=summary['cold_result'], dll_sha256=sha((folder/'rpfem_native_x64.dll').read_bytes()),
                  native_analysis_seconds=summary['analysis_seconds'], frozen_native_module_count=len(summary['source_sha256']),
                  native_summary_literal_proof_count=len(summary['independent_raw_replay']),
                  audit_script_sha256=sha(Path(__file__).read_bytes()), original_python_source_hashes={str(p.relative_to(REPO)):sha(p.read_bytes()) for p in sorted(PYTHON.rglob('*.py')) if '__pycache__' not in p.parts})
    core.write(args.output, output)
    print(json.dumps({k:output[k] for k in ('status','lower','upper','gap_percent','literal_original_proof_count')}, ensure_ascii=False))


if __name__ == '__main__':
    main()
