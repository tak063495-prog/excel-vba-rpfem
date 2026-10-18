from pathlib import Path
import sys,struct,math,json
import numpy as np
from prepare import REPO
sys.path.insert(0,str(REPO/'versions/G1R13/validation_runtime/python'))
from common import Mesh
from independent_audit import audit
from zero_phi import audit_zero
def read_state(path,policy,phase):
    raw=path.read_bytes();assert raw[:8]==b'RP25M001';pos=8
    def take(dtype,count):
        nonlocal pos
        a=np.frombuffer(raw,dtype=dtype,count=count,offset=pos).copy();pos+=a.nbytes;return a
    nn,ne,nm,nr,q=map(int,take('<i4',5));assert nr==0
    scale,factor,fs,objective,native_audit=take('<f8',5)
    xy=take('<f8',nn*2).reshape(nn,2);tri=take('<i4',ne*5).reshape(ne,5)
    materials=take('<f8',nm*4).reshape(nm,4);body=take('<f8',ne*4).reshape(ne,2,2)
    nb=int(take('<i4',1)[0]);boundary={};kinds=['load','fixed','roller_x','roller_y']
    for _ in range(nb):
        a,b,k,r=map(int,take('<i4',4));loads=take('<f8',4).reshape(2,2).T
        boundary[tuple(sorted((a,b))) ]={'kind':kinds[k],'rigid':r,'loads':loads}
    nf=int(take('<i4',1)[0]);x=take('<f8',nf);assert pos==len(raw)
    assert phase=='reference_total' and factor==fs and np.all(tri[:,4]<0)
    m=Mesh(xy,tri[:,:3],materials,tri[:,3],tri[:,4],body,boundary,float(scale),'native_dll2_26',
           load_phase=phase,provenance={'psi_policy':policy}).validate()
    return m,x,{'nodes':nn,'elements':ne,'q':q,'factor':factor,'fs':fs,'objective':objective,'native_audit':native_audit,'scale':scale}
def expected_body(m,water):
    body=np.zeros_like(m.body);body[:,1,1]=-m.materials[m.material_id,3]
    if water['coupling']:
        # Independent P1 interpolation of the declared nodal hydrostatic pressure.
        p=m.xy[m.tri];u=water['gamma_w']*np.maximum(3.-p[:,:,1],0.)
        a=p[:,1]-p[:,0];b=p[:,2]-p[:,0];det=a[:,0]*b[:,1]-a[:,1]*b[:,0]
        du1=u[:,1]-u[:,0];du2=u[:,2]-u[:,0]
        grad=np.stack(((du1*b[:,1]-du2*a[:,1])/det,(a[:,0]*du2-b[:,0]*du1)/det),axis=1)
        body[:,:,1]-=grad
    return body
def audit_case(case,folder,rec,native):
    c=rec['source_case'];checks=[]
    for index,side in enumerate(['lower','upper']):
        m,x,meta=read_state(folder/f'witness_{side}.bin',c['policy'],native['witness_phase'][index])
        assert meta['q']==native['witness_q'][index]
        assert math.isclose(meta['fs'],native[side+'_fs'],rel_tol=1e-12)
        wanted=np.array([[v['c'],v['phi'],v['psi'],v['gamma']] for v in c['materials']])
        assert np.array_equal(m.materials,wanted)
        p=m.xy[m.tri];a=p[:,1]-p[:,0];b=p[:,2]-p[:,0];areas=(a[:,0]*b[:,1]-a[:,1]*b[:,0])/2
        by_material=[float(areas[m.material_id==j].sum()) for j in range(len(wanted))]
        assert np.allclose(by_material,rec['region_areas'],rtol=1e-9,atol=1e-8)
        predicted=expected_body(m,rec['water']);body_error=float(np.max(np.abs(predicted-m.body)))
        assert body_error<=1e-9,(case,side,'water/gravity mismatch',body_error)
        actual_weight=float(-areas@m.body[:,1,1])
        original_weight=float(areas@m.materials[m.material_id,3]);assert math.isclose(original_weight,rec['total_weight'],rel_tol=1e-9)
        tolerance=1e-10*max(np.ptp(m.xy,axis=0).max(),1.)
        def on_segment(point,start,end):
            start=np.array(start);end=np.array(end);d=end-start;v=point-start
            t=float(v@d/(d@d));return -1e-10<=t<=1+1e-10 and abs(d[0]*v[1]-d[1]*v[0])<=tolerance*np.linalg.norm(d)
        kinds={'自由':'load','固定XY':'fixed','固定X':'roller_x','固定Y':'roller_y'}
        for edge_key,edge in m.boundary.items():
            ea,eb=m.xy[list(edge_key)]
            matches=[v for v in rec['boundary_explicit'] if on_segment(ea,v['a'],v['b']) and on_segment(eb,v['a'],v['b'])]
            assert len(matches)==1,(case,edge_key)
            desired=matches[0];assert edge['kind']==kinds[desired['kind']]
            assert np.array_equal(edge['loads'][:,0],np.zeros(2))
            assert np.array_equal(edge['loads'][:,1],np.array(desired['load0'])+np.array(desired['load1']))
        result=audit_zero(m,x,meta['fs'],meta['q'],meta['objective']) if side=='upper' and np.all(m.materials[:,1]==0) else audit(m,x,side,meta['fs'],meta['q'],meta['objective'],18*len(m.tri))
        raw={k:max(0.,float(v))*m.stress_scale for k,v in result['errors'].items()} if side=='lower' else {}
        raw_ok=max(raw.values(),default=0.)<=1e-7
        endpoint=float(x[18*len(m.tri)])>=1.-1e-8 if side=='lower' else result['value']<=1.+1e-8
        checks.append({'side':side,'metadata':meta,'physical_audit':result,'raw_lower_errors':raw,
            'raw_lower_pass':bool(raw_ok),'endpoint_direction_pass':bool(endpoint),
            'material_areas':by_material,'gravity_weight':original_weight,'effective_vertical_load':actual_weight,
            'independent_body_max_error':body_error,'pore_pressure_on':bool(rec['water']['coupling']),
            'pass':bool(result['pass'] and raw_ok and endpoint)})
    return {'pass':all(c['pass'] for c in checks),'scope':'Independent physical equations, material/geometry, explicit boundaries, dry weight and D08A P1 water-gradient loads','results':checks}
