"""Emit RP25G001 from xy/tri/material only; never reads a field/Fs/root file."""
from pathlib import Path
import argparse,hashlib,json,os,struct,sys
for name in ['OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS']:os.environ[name]='1'
REPO=Path(__file__).resolve().parents[3]
sys.path.insert(0,str(REPO/'validation/G1R19_python26_20261016/python'))
import benchmark_models
import experiment_core as core
np=core.np
def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()
def export(mesh_path,case,output,expected_sha):
    mesh_path=Path(mesh_path).resolve();output=Path(output).resolve()
    if output.exists():raise FileExistsError('Refusing geometry overwrite')
    assert sha(mesh_path)==expected_sha
    canonical=REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs'/(case+'.json')
    record=benchmark_models._record(case)
    # np.load is lazy. Read only the three geometry arrays, irrespective of any
    # extra archive metadata. No field/certificate paths are opened.
    with np.load(mesh_path,allow_pickle=False) as data:
        xy=data['xy'].copy();tri=data['tri'].copy();material_id=data['material_id'].copy();keys=list(data.files)
    mesh=core.mesh_from_arrays(xy,tri,material_id,record)
    assert len(mesh.materials)>0 and np.all(mesh.rigid_id==-1)
    cells=np.column_stack((mesh.tri,mesh.material_id,mesh.rigid_id)).astype('<i4')
    pieces=[b'RP25G001',struct.pack('<5i',len(mesh.xy),len(mesh.tri),len(mesh.materials),0,8),
            struct.pack('<d',mesh.stress_scale),np.asarray(mesh.xy,'<f8').tobytes(),cells.tobytes(),
            np.asarray(mesh.materials,'<f8').tobytes(),np.asarray(mesh.body,'<f8').tobytes(),struct.pack('<i',len(mesh.boundary))]
    for (a,b),boundary in mesh.boundary.items():
        kind=['load','fixed','roller_x','roller_y'].index(boundary['kind'])
        pieces+=[struct.pack('<4i',a,b,kind,boundary['rigid']),np.asarray(boundary['loads'].T,'<f8').tobytes()]
    raw=b''.join(pieces);output.parent.mkdir(parents=True,exist_ok=True);output.write_bytes(raw)
    # The strict geometry reader checks EOF here: additional field/Fs bytes are
    # a format error, even if a future exporter accidentally appends them.
    import importlib.util
    spec=importlib.util.spec_from_file_location('geometry_reader',REPO/'versions/G1R19_2pct/tests/run_mesh_export_excel.py')
    reader=importlib.util.module_from_spec(spec);spec.loader.exec_module(reader)
    read=reader.read_geometry(output)
    for a,b in [(read[0],mesh.xy),(read[1],cells),(read[2],mesh.materials),(read[3],mesh.body)]:np.testing.assert_array_equal(a,b)
    assert read[4].keys()==mesh.boundary.keys()
    for edge,got in read[4].items():
        expected=mesh.boundary[edge]
        assert got['kind']==expected['kind'] and got['rigid']==expected['rigid']
        np.testing.assert_array_equal(got['loads'],expected['loads'])
    result={'status':'PASS','scope':'Known-geometry replay entry, not automatic cold meshing',
        'case':case,'nodes':len(mesh.xy),'elements':len(mesh.tri),'q_for_new_solve':8,
        'mesh_source':str(mesh_path),'mesh_source_sha256':expected_sha,'archive_keys':keys,'read_arrays':['xy','tri','material_id'],
        'canonical_input':str(canonical),'canonical_input_sha256':sha(canonical),'physical_model_id':mesh.physical_model_id,
        'geometry_id':mesh.mesh_id,'source_field_opened':False,'Fs_root_status_imported':False,'solver_calls':0,
        'binary':str(output),'binary_sha256':sha(output),'binary_bytes':len(raw),
        'exporter_sha256':sha(__file__),'python_source_sha256':{p.name:sha(p) for p in [Path(core.__file__),Path(benchmark_models.__file__)]}}
    output.with_suffix('.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf8')
    return result
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--mesh',type=Path,required=True);p.add_argument('--case',choices=['P06A','P06N'],required=True);p.add_argument('--sha256',required=True);p.add_argument('--output',type=Path,required=True)
    a=p.parse_args();print(json.dumps(export(a.mesh,a.case,a.output,a.sha256)))
