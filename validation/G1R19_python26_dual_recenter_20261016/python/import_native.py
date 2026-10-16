"""Re-audit native26 binary witnesses against the explicit benchmark models.

Source labels, solver status, and native root flags never establish a certificate.
Only the original physical audit establishes an endpoint; root searches are new.
"""
from pathlib import Path
import argparse
import hashlib
import json
import zipfile
import accuracy_driver as driver
import experiment_core as core
from experiment_core import np, Prepared, evaluate, write
from model_support.upper_assembly import complete_upper_velocity


def read_state(raw):
    if raw[:8] != b'RP25M001':
        raise ValueError('Unsupported native witness format')
    position = 8

    def take(dtype, count):
        nonlocal position
        if count < 0 or position + np.dtype(dtype).itemsize * count > len(raw):
            raise ValueError('Truncated native witness')
        value = np.frombuffer(raw, dtype=dtype, count=count, offset=position).copy()
        position += value.nbytes
        return value

    nn, ne, nm, nr, q = map(int, take('<i4', 5))
    if min(nn, ne, nm, q) <= 0 or nr != 0:
        raise ValueError('Only soil-only witnesses with positive quadrature are supported')
    scale, factor, fs, objective, native_audit = map(float, take('<f8', 5))
    xy = take('<f8', 2*nn).reshape(nn, 2)
    cells = take('<i4', 5*ne).reshape(ne, 5)
    materials = take('<f8', 4*nm).reshape(nm, 4)
    body = take('<f8', 4*ne).reshape(ne, 2, 2)
    boundary = {}
    kinds = ['load', 'fixed', 'roller_x', 'roller_y']
    for _ in range(int(take('<i4', 1)[0])):
        a, b, kind, rigid = map(int, take('<i4', 4))
        if not 0 <= kind < len(kinds) or (min(a,b) < 0 or max(a,b) >= nn):
            raise ValueError('Invalid boundary')
        key = tuple(sorted((a,b)))
        if key in boundary:
            raise ValueError('Duplicate boundary')
        boundary[key] = {'kind': kinds[kind], 'rigid': rigid,
                         'loads': take('<f8', 4).reshape(2,2).T}
    field = take('<f8', int(take('<i4',1)[0]))
    if position != len(raw) or factor != fs or np.any(cells[:,4] != -1):
        raise ValueError('Unsupported phase, rigid body, or trailing bytes')
    if not driver.finite(scale, fs, objective, native_audit) or not fs > 0:
        raise ValueError('Nonfinite native metadata')
    return xy, cells, materials, body, boundary, field, {'q':q, 'fs':fs, 'scale':scale,
              'objective':objective, 'source_sha256':hashlib.sha256(raw).hexdigest()}


def import_side(raw, record, rec, mode):
    xy, cells, materials, body, boundary, field, metadata = read_state(raw)
    mesh = core.mesh_from_arrays(xy, cells[:,:3], cells[:,3], rec)
    if not np.array_equal(materials,mesh.materials) or not np.array_equal(body,mesh.body):
        raise ValueError('Native material or body load differs from benchmark')
    if metadata['scale'] != mesh.stress_scale or boundary.keys() != mesh.boundary.keys():
        raise ValueError('Native scale or boundary differs from benchmark')
    for edge, expected in mesh.boundary.items():
        got = boundary[edge]
        if got['kind'] != expected['kind'] or got['rigid'] != expected['rigid'] or not np.array_equal(got['loads'],expected['loads']):
            raise ValueError('Native boundary load or restraint differs from benchmark')
    index = ['lower','upper'].index(mode)
    if record['witness_phase'][index] != 'reference_total' or record['witness_q'][index] != metadata['q'] or record[mode+'_fs'] != metadata['fs']:
        raise ValueError('Native manifest and witness differ')
    fs = metadata['fs']
    if mode == 'upper':
        # Rebuild the current conservative phi=0 epigraph from velocity alone.
        # Old auxiliary prices cannot establish a new physical upper bound.
        field = complete_upper_velocity(mesh,field[:12*len(mesh.tri)],8)
    original = evaluate(Prepared(mesh,fs,mode,8),field)
    attempts = [{'fs':fs,'audit':original}]
    if not original['witness'] and original['physical_pass']:
        for backoff in [1e-5,1e-4,1e-3]:
            newfs = fs*(1-backoff if mode == 'lower' else 1+backoff)
            candidate = field.copy()
            if mode == 'lower':
                if original['value'] <= 0 or np.any(mesh.body[:,:,0]) or any(np.any(v['loads'][:,0]) for v in mesh.boundary.values()):
                    break
                candidate *= 1.0000012/original['value']
            result = evaluate(Prepared(mesh,newfs,mode,8),candidate)
            attempts.append({'fs':newfs,'audit':result})
            if result['witness']:
                fs, field, original = newfs,candidate,result
                break
    if not original['witness']:
        return None, {'accepted':False,'source':metadata,'attempts':attempts}
    original = dict(original,native_import={'source':metadata,'attempts':attempts,
                                            'root_flags_reused':False})
    certificate = driver.Certificate(mode,fs,mesh,field,original,'native26 independent re-audit',False)
    return certificate, {'accepted':True,'source':metadata,'attempts':attempts}


def run(archive, output, cases):
    output.mkdir(parents=True,exist_ok=True)
    reports = []
    with zipfile.ZipFile(archive) as source:
        for case in cases:
            (output/case).mkdir(parents=True,exist_ok=True)
            _,rec=driver.load_model(case)
            record=json.loads(source.read(f'G1R17_native26/results/{case}.json'))
            certificates={}
            report={'case':case,'sides':{}}
            for mode in ['lower','upper']:
                try:
                    cert, audit=import_side(source.read(f'G1R17_native26/cases/{case}/witness_{mode}.bin'),record,rec,mode)
                except ValueError as error:
                    cert,audit=None,{'accepted':False,'reason':str(error)}
                report['sides'][mode]=audit
                if cert is not None:
                    certificates[mode]=cert.publish(output/case)
            write(output/case/'import_audit.json',{'case':case,'certificates':certificates,
                 'scope':'Native26 original physical re-audit; no inherited root completion'})
            reports.append(report)
            write(output/'native_import_report.json',reports)
            print('IMPORTED_NATIVE',case,{mode:c['fs'] for mode,c in certificates.items()},flush=True)


if __name__ == '__main__':
    parser=argparse.ArgumentParser()
    parser.add_argument('--archive',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--cases',nargs='+',required=True)
    args=parser.parse_args()
    run(args.archive.resolve(),args.output.resolve(),args.cases)
