"""Independent original-matrix replay of VBA dual and negative primal evidence."""
from pathlib import Path
import json
import math
import struct


class Reader:
    def __init__(self, path, magic):
        self.raw = Path(path).read_bytes()
        assert self.raw[:8] == magic, 'Wrong evidence magic'
        self.position = 8

    def read(self, format):
        layout = struct.Struct('<' + format)
        result = layout.unpack_from(self.raw, self.position)
        self.position += layout.size
        return result[0] if len(result) == 1 else result

    def finish(self):
        assert self.position == len(self.raw), 'Truncated or trailing evidence'


def context(path):
    text = Path(path).read_bytes().decode('utf-16')
    result = {}
    for line in text.splitlines():
        parts = line.split('\t')
        result.setdefault(parts[0], []).append(parts[1:])
    return result


def exact_float(hex_value):
    return struct.unpack('<d', bytes.fromhex(hex_value))[0]


def replay_dual(path):
    reader = Reader(path, b'RG18F001')
    n, m, blocks, s, q = reader.read('5i')
    fs, recorded_stat, recorded_cone, recorded_sep = reader.read('4d')
    assert min(n, m, blocks, s) > 0 and q == 8 and fs > 0
    costs = [reader.read('d') for _ in range(n)]
    assert all(value == 0 for value in costs), 'Nonzero original objective'
    rows = []
    for _ in range(m):
        rhs, norm, normalized_y, raw_y = reader.read('4d')
        assert norm > 0 and raw_y == normalized_y / norm
        count = reader.read('i')
        assert 0 < count <= n
        values = [reader.read('id') for _ in range(count)]
        assert all(0 <= index < n for index, _ in values)
        rows.append((rhs, norm, normalized_y, raw_y, values))
    cones = []
    total = 0
    for _ in range(blocks):
        dim, count, first = reader.read('3i')
        assert dim in (1, 3) and 0 < count <= n and first == total
        indices = [reader.read('i') for _ in range(count)]
        assert all(0 <= index < n for index in indices)
        matrix = [(reader.read('d'), [reader.read('d') for _ in range(count)]) for _ in range(dim)]
        cones.append((dim, indices, matrix))
        total += dim
    assert total == s
    z = [reader.read('d') for _ in range(s)]
    reader.finish()
    stationarity = [[] for _ in range(n)]
    separation = []
    numeric = [fs, *costs, *z]
    for rhs, norm, normalized_y, raw_y, values in rows:
        numeric += [rhs, norm, normalized_y, raw_y, *(value for _, value in values)]
        separation.append(-rhs * raw_y)
        for index, value in values:
            stationarity[index].append(value * raw_y)
    cone_error, first = 0., 0
    for dim, indices, matrix in cones:
        for offset, coefficients in matrix:
            numeric += [offset, *coefficients]
            separation.append(-offset * z[first])
            for index, value in zip(indices, coefficients):
                stationarity[index].append(-value * z[first])
            first += 1
        head = z[first - dim]
        error = -head if dim == 1 else math.hypot(z[first - 2], z[first - 1]) - head
        cone_error = max(cone_error, error)
    assert all(math.isfinite(value) for value in numeric)
    stat = max(abs(math.fsum(values)) for values in stationarity)
    sep = math.fsum(separation)
    assert stat <= 1e-7 and cone_error <= 1e-8 and sep >= 1.000001
    assert abs(sep - 1.000003) <= 1e-7
    for value, recorded in [(stat, recorded_stat), (cone_error, recorded_cone), (sep, recorded_sep)]:
        assert abs(value - recorded) <= 2e-12 * (1 + abs(recorded))
    ctx = context(Path(path).with_suffix('.context.tsv'))
    assert exact_float(ctx['fs'][0][0]) == fs and ctx['q'] == [['8']]
    assert 'load_phase=reference_total' in next(iter(ctx))
    return {'path': str(path), 'fs': fs, 'q': q, 'stationarity': stat,
            'cone': cone_error, 'separation': sep, 'positive_guide_only': True}


def replay_negative(prefix):
    prefix = Path(prefix)
    reader = Reader(str(prefix) + '.negative_program.bin', b'P06VBA01')
    n, m, blocks = reader.read('3i')
    costs = [reader.read('d') for _ in range(n)]
    rows = []
    for _ in range(m):
        count, rhs = reader.read('id')
        rows.append((rhs, [reader.read('id') for _ in range(count)]))
    cones = []
    for _ in range(blocks):
        dim = reader.read('i')
        assert dim in (1, 3)
        matrix = []
        for _ in range(dim):
            count, offset = reader.read('id')
            matrix.append((offset, [reader.read('id') for _ in range(count)]))
        cones.append(matrix)
    reader.finish()
    field_reader = Reader(str(prefix) + '.negative_field.bin', b'P06X0001')
    assert field_reader.read('i') == n
    x = [field_reader.read('d') for _ in range(n)]
    field_reader.finish()
    ctx = context(str(prefix) + '.negative_context.tsv')
    norms = {int(index): exact_float(value) for index, value in ctx['row_norm']}
    fs = exact_float(ctx['fs'][0][0])
    assert ctx['q'] == [['8']]
    assert n > 0 and m > 0 and blocks > 0 and fs > 0 and math.isfinite(fs)
    assert all(math.isfinite(value) for value in [*costs, *x, *norms.values()])
    assert all(norms[i] > 0 for i in range(m))
    for rhs, values in rows:
        assert math.isfinite(rhs) and all(0 <= index < n and math.isfinite(value) for index, value in values)
    for matrix in cones:
        for offset, values in matrix:
            assert math.isfinite(offset) and all(0 <= index < n and math.isfinite(value) for index, value in values)
    raw_eq = max(abs(math.fsum([*(value * x[index] for index, value in values), -rhs])) * norms[i]
                 for i, (rhs, values) in enumerate(rows))
    cone_error = 0.
    for matrix in cones:
        values = [math.fsum([offset, *(value * x[index] for index, value in coefficients)])
                  for offset, coefficients in matrix]
        cone_error = max(cone_error, -values[0] if len(values) == 1 else math.hypot(*values[1:]) - values[0])
    objective = math.fsum(c * value for c, value in zip(costs, x))
    assert raw_eq <= 1e-7 and cone_error <= 1e-7 and objective <= .999999 + 1e-8
    return {'path': str(prefix), 'fs': fs, 'q': 8, 'raw_equality': raw_eq,
            'raw_cone': cone_error, 'objective': objective, 'physical_gate_executed_in_vba': True}


def verify_folder(folder):
    folder = Path(folder)
    paths = list((folder / 'RPFEM_logs').glob('*.original_guide_*.bin'))
    assert paths, 'No retained original ray vectors/program'
    duals = [replay_dual(path) for path in paths]
    negative = replay_negative(folder / 'c0_wedge_native.tsv')
    positive = max(duals, key=lambda value: value['fs'])
    assert positive['fs'] < negative['fs']
    width = negative['fs'] - positive['fs']
    assert width <= 1e-4 * max(1., (negative['fs'] + positive['fs']) / 2)
    report = {'status': 'PASS', 'original_dual_proofs': duals,
              'negative_original_primal': negative, 'root_width': width,
              'scope': 'Independent matrix replay; raw ray is only positive discrete guide'}
    (folder / 'original_evidence_replay.json').write_text(json.dumps(report, indent=2) + '\n', encoding='utf-8')
    return report


if __name__ == '__main__':
    import sys
    print(json.dumps(verify_folder(sys.argv[1]), indent=2))
