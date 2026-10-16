"""Immutable geometry-only interchange fixtures; no optimization or root reuse."""
from pathlib import Path
import argparse, ast, copy, hashlib, inspect, json, os, struct, sys
for name in ['OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS']:os.environ[name]='1'
ROOT=Path(__file__).resolve().parents[1]
REPO=ROOT.parents[1]
PYTHON=REPO/'validation/G1R19_python26_20261016/python'
sys.path.insert(0,str(PYTHON))
import experiment_core as core
import mesh_geometry as reference
import benchmark_models
np=core.np


def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()


def protected_reference(mesh,record,velocity,extra=(),passes=5):
    """Exact referenced Python algorithm with one additional guard predicate.

    The original function protects exterior/material boundaries intrinsically.
    The VBA API additionally accepts existing RConstraints. Only that optional
    membership rejection is inserted; metrics/ordering/thresholds are unchanged.
    """
    tree=ast.parse(inspect.getsource(reference.flow_metric_flip))
    nodes=[n for n in ast.walk(tree) if isinstance(n,ast.If) and ast.unparse(n.test)==
           'f < 0 or mesh.material_id[e] != mesh.material_id[f]']
    if len(nodes)!=1:raise ValueError('Reference protection anchor changed')
    nodes[0].test=ast.BoolOp(op=ast.Or(),values=[nodes[0].test,
        ast.Compare(left=ast.Tuple(elts=[ast.Attribute(value=ast.Name(id='edge',ctx=ast.Load()),attr='a',ctx=ast.Load()),
                                        ast.Attribute(value=ast.Name(id='edge',ctx=ast.Load()),attr='b',ctx=ast.Load())],ctx=ast.Load()),
                    ops=[ast.In()],comparators=[ast.Name(id='EXTRA_PROTECTED',ctx=ast.Load())])])
    ast.fix_missing_locations(tree)
    namespace=reference.flow_metric_flip.__globals__.copy();namespace['EXTRA_PROTECTED']=set(extra)
    exec(compile(tree,'exact_python_flow_with_explicit_constraints','exec'),namespace)
    return namespace['flow_metric_flip'](mesh,record,velocity,passes)


def write_input(path,mesh,velocity,protected,passes):
    guards=np.array(sorted(protected),np.int32).reshape(-1,2)
    with path.open('wb') as stream:
        stream.write(b'FMG1I001');stream.write(struct.pack('<4i',len(mesh.xy),len(mesh.tri),len(guards),passes))
        for value in [np.asarray(mesh.xy,'<f8'),np.asarray(mesh.tri,'<i4'),np.asarray(mesh.material_id,'<i4'),
                      np.asarray(velocity,'<f8'),np.asarray(guards,'<i4')]:stream.write(value.tobytes())


def read_input(path):
    raw=Path(path).read_bytes();offset=24
    if raw[:8]!=b'FMG1I001':raise ValueError('Wrong input magic')
    nn,ne,ng,passes=struct.unpack_from('<4i',raw,8)
    def take(dtype,count):
        nonlocal offset
        value=np.frombuffer(raw,dtype=dtype,count=count,offset=offset).copy();offset+=value.nbytes;return value
    xy=take('<f8',2*nn).reshape(nn,2);tri=take('<i4',3*ne).reshape(ne,3)
    mat=take('<i4',ne);velocity=take('<f8',12*ne);guards=take('<i4',2*ng).reshape(ng,2)
    if offset!=len(raw):raise ValueError('Input bytes not fully consumed')
    return xy,tri,mat,velocity,guards,passes


def read_output(path):
    raw=Path(path).read_bytes()
    if raw[:8]!=b'FMG1O001':raise ValueError('Wrong output magic')
    nn,ne,nflip,certificate=struct.unpack_from('<4i',raw,8);seconds=struct.unpack_from('<d',raw,24)[0]
    tri=np.frombuffer(raw,'<i4',3*ne,offset=32).copy().reshape(ne,3);offset=32+12*ne
    changes=[]
    for _ in range(nflip):
        values=struct.unpack_from('<7id',raw,offset);offset+=36
        changes.append(dict(old=list(values[:2]),new=list(values[2:4]),elements=list(values[4:6]),sweep=values[6],gain=values[7]))
    if offset!=len(raw) or certificate!=0:raise ValueError('Trailing output bytes/certificate flag')
    return tri,dict(nodes=nn,elements=ne,flips=nflip,certificate=False,seconds=seconds,changes=changes)


def make(output):
    if output.exists():raise FileExistsError('Refusing to replace fixture directory')
    output.mkdir(parents=True)
    benchmark_models.install_model_builder()
    entries=[]
    for case,folder in [('P06A','a_flowflip/P06A/source_witness'),('P06N','n_flow_geometry/P06N/source_witness')]:
        root=REPO/'validation/accuracy_20261013/results'/folder
        record=benchmark_models._record(case)
        with np.load(root/'best_upper_mesh.npz',allow_pickle=False) as data:
            mesh=core.mesh_from_arrays(data['xy'],data['tri'],data['material_id'],record)
        with np.load(root/'best_upper_field.npz',allow_pickle=False) as data:
            if str(data['mode'])!='upper' or int(data['q'])!=8:raise ValueError('Source is not upper q8')
            velocity=data['x'][:12*len(mesh.tri)].copy();fs=float(data['fs'])
        source_audit=core.evaluate(core.Prepared(mesh,fs,'upper',8,reduction='none'),velocity)
        if not source_audit['physical_pass'] or not source_audit['witness']:raise ValueError('Archived source upper field failed physical audit')
        for explicit in ([False,True] if case=='P06N' else [False]):
            extra=set()
            if explicit:
                _,plain=reference.flow_metric_flip(mesh,record,velocity,passes=1)
                extra={tuple(sorted(c['old'])) for c in plain['changes'][:8]}
            child,info=protected_reference(mesh,record,velocity,extra,5)
            guards=reference.protected_edges(mesh)|extra
            name=case+'_actual'+str(len(mesh.tri))+('_constraints' if explicit else '')
            write_input(output/(name+'.bin'),mesh,velocity,guards,5)
            np.savez_compressed(output/(name+'_expected.npz'),tri=child.tri)
            (output/(name+'_expected.json')).write_text(json.dumps(info,indent=2)+'\n',encoding='utf8')
            entries.append(dict(name=name,case=case,input=name+'.bin',input_sha256=sha(output/(name+'.bin')),
                                expected_tri=name+'_expected.npz',expected_tri_sha256=sha(output/(name+'_expected.npz')),
                                expected_info=name+'_expected.json',expected_info_sha256=sha(output/(name+'_expected.json')),
                                elements=len(mesh.tri),protected_edges=len(guards),extra_constraints=len(extra),flips=info['flips'],
                                source_files={p.relative_to(REPO).as_posix():sha(p) for p in [root/'best_upper_mesh.npz',root/'best_upper_field.npz']},
                                source_fs=fs,source_physical_audit=source_audit,
                                scope='Archived source velocity used for geometry comparison only; candidate has no field/root/bound certificate'))
            print(name,len(mesh.tri),info['flips'],flush=True)
    data=dict(status='PYTHON_EXPECTED_GEOMETRY_ONLY; VBA_RUNTIME_PENDING',solver_calls=0,
              reference=reference.__file__,reference_sha256=sha(Path(reference.__file__)),
              explicit_constraints_rule='Only an extra source-edge rejection predicate added to original Python function',fixtures=entries)
    (output/'FIXTURES.json').write_text(json.dumps(data,indent=2)+'\n',encoding='utf8')
    return data


if __name__=='__main__':
    parser=argparse.ArgumentParser(description=__doc__);parser.add_argument('--output',type=Path,required=True)
    args=parser.parse_args();result=make(args.output)
    print(json.dumps(dict(status=result['status'],fixtures=len(result['fixtures']))))
