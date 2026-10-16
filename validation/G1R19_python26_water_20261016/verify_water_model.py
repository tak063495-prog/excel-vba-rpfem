"""Independent geometry and common-load checks; no numerical optimization."""
from pathlib import Path
import hashlib,json,sys
import numpy as np

ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'python'))
import benchmark_models as models
import experiment_core as core


def sha(path):return hashlib.sha256(path.read_bytes()).hexdigest()


def polygon_moments(points):
    p=np.asarray(points,float);q=np.roll(p,-1,axis=0)
    cross=p[:,0]*q[:,1]-q[:,0]*p[:,1]
    return np.array([cross.sum()/2,((p[:,0]+q[:,0])*cross).sum()/6,
                     ((p[:,1]+q[:,1])*cross).sum()/6])


def inspect(mesh,record,reference=None):
    area,grad,edges=core.topology(mesh);points=mesh.xy[mesh.tri];y=points[:,:,1]
    crossing=(y.min(axis=1)<3)&(y.max(axis=1)>3)
    if crossing.any():raise ValueError('A triangle straddles the physical pressure kink')
    below=y.mean(axis=1)<3
    gamma=mesh.materials[mesh.material_id,3]
    if not (np.all(mesh.body[:,:,0]==0) and np.all(mesh.body[:,0,1]==0) and
            np.array_equal(mesh.body[:,1,1],-gamma+9.81*below)):
        raise ValueError('The common continuum hydrostatic body is not exact')
    moments=np.column_stack((area,area[:,None]*points.mean(axis=1))).sum(axis=0)
    wanted=polygon_moments(record['source_case']['outline'])
    if not np.allclose(moments,wanted,rtol=0,atol=1e-10):raise ValueError('Domain moments changed')
    # The prescribed benchmark has the full 20x3 rectangle below the line.
    wet=np.column_stack((area[below],area[below,None]*points[below].mean(axis=1))).sum(axis=0)
    if not np.allclose(wet,[60.,600.,90.],rtol=0,atol=1e-10):raise ValueError('Wet-domain measure changed')
    body_moments=np.array([np.dot(area,mesh.body[:,1,1]),
                          np.dot(area*points[:,:,0].mean(axis=1),mesh.body[:,1,1]),
                          np.dot(area*points[:,:,1].mean(axis=1),mesh.body[:,1,1])])
    expected=-20*wanted+9.81*np.array([60.,600.,90.])
    if not np.allclose(body_moments,expected,rtol=0,atol=1e-9):raise ValueError('Force or first moment differs from prescribed body')
    # On these conforming triangles independently differentiated nodal P1 u
    # must equal the explicit continuum gradient to floating roundoff.
    pressure=9.81*np.maximum(3-y,0)
    p1_gradient=np.einsum('ek,ekd->ed',pressure,grad)
    exact=np.column_stack((np.zeros(len(area)),-9.81*below))
    gradient_error=float(np.max(abs(p1_gradient-exact)))
    if gradient_error>1e-10:raise ValueError('P1 water field does not match the continuum gradient')
    boundary={}
    for edge in edges:
        if edge.elements[1]<0:
            boundary[edge.kind]=boundary.get(edge.kind,0.)+edge.length
            if np.any(edge.loads):raise ValueError('D08A should have no external water traction')
    parent_body_exact=None
    if reference is not None:
        parent=np.asarray(mesh.provenance['parent_element_ids'],int)
        parent_body_exact=np.array_equal(mesh.body,reference.body[parent])
        if not parent_body_exact:raise ValueError('Refinement changed the represented body field')
        if mesh.physical_model_id!=reference.physical_model_id:raise ValueError('Physical identity changed')
    return {'elements':len(mesh.tri),'nodes':len(mesh.xy),'mesh_id':mesh.mesh_id,
            'physical_model_id':mesh.physical_model_id,'waterline_crossing_triangles':0,
            'domain_area_first_moments':moments.tolist(),'wet_area_first_moments':wet.tolist(),
            'vertical_body_force_first_moments':body_moments.tolist(),
            'horizontal_body_force_exact_zero':True,'boundary_lengths':boundary,
            'nodal_P1_gradient_max_error':gradient_error,'parent_body_preserved_exactly':parent_body_exact}


def verify():
    models.install_model_builder();mesh,record=models.load_model('D08A')
    initial=inspect(mesh,record)
    uniform=core.uniform_refine(mesh,record);uniform_result=inspect(uniform,record,mesh)
    local,change=core.local_refine(mesh,record,np.arange(1,len(mesh.tri)+1,dtype=float),.35)
    local_result=inspect(local,record,mesh)
    for result in (uniform_result,local_result):
        for kind,length in initial['boundary_lengths'].items():
            if abs(result['boundary_lengths'].get(kind,0)-length)>1e-10:raise ValueError('Boundary condition lengths changed')
    return {'status':'PASS','solver_calls':0,'scope':'Exact common prescribed hydrostatic load on conforming initial and nested refined meshes',
            'old_saved_fields_imported':False,'old_native_roots_imported':False,
            'initial':initial,'uniform':uniform_result,'local':local_result,
            'manifest_sha256':sha(ROOT/'INPUTS_26.json'),
            'implementation_sha256':{p.relative_to(ROOT).as_posix():sha(p) for p in sorted((ROOT/'python').rglob('*.py'))}}


if __name__=='__main__':
    result=verify();(ROOT/'MODEL_VERIFICATION.json').write_text(json.dumps(result,indent=2,ensure_ascii=False)+'\n',encoding='utf8')
    print(json.dumps({key:result[key] for key in ('status','solver_calls','initial','uniform','local')},ensure_ascii=False),flush=True)
