"""Compare interpolation arithmetic while keeping the original matrix/audit gates."""
from pathlib import Path
import argparse
import json
import time
from scipy.spatial import cKDTree

import accuracy_driver as driver
from experiment_core import np,Prepared,evaluate,topology,write
from model_support.p2 import gradients
from independent_audit import strength
from mesh_geometry import interpolate_velocity
from precision_interpolation import precision_prolong_upper


def local_bary(vertices,points):
    vertices=np.asarray(vertices,dtype=np.longdouble);points=np.asarray(points,dtype=np.longdouble)
    a=vertices[...,0,:];u=vertices[...,1,:]-a;v=vertices[...,2,:]-a;w=points-a
    determinant=u[...,0]*v[...,1]-u[...,1]*v[...,0]
    b=(w[...,0]*v[...,1]-w[...,1]*v[...,0])/determinant
    c=(u[...,0]*w[...,1]-u[...,1]*w[...,0])/determinant
    return np.stack([1-b-c,b,c],axis=-1)


def nested_longdouble_interpolation(parent,child,field):
    return precision_prolong_upper(parent,child,field,8,return_parent_ids=True)


def legacy_global_inverse_interpolation(parent,child,field):
    """Frozen old diagnostic arithmetic; never used for publishing a bound."""
    source=field[:12*len(parent.tri)].reshape(-1,6,2);vertices=parent.xy[parent.tri]
    matrix=np.concatenate([vertices.transpose(0,2,1),np.ones((len(vertices),1,3))],axis=1)
    inverse=np.linalg.inv(matrix);tree=cKDTree(vertices.mean(axis=1))
    target=child.xy[child.tri];points=np.concatenate([target,(target+np.roll(target,-1,axis=1))/2],axis=1)
    centers=target.mean(axis=1);_,nearby=tree.query(points.reshape(-1,2),k=min(32,len(vertices)))
    nearby=np.asarray(nearby).reshape(len(child.tri),6,-1);output=np.empty((len(child.tri),6,2))
    for e in range(len(child.tri)):
        for k,point in enumerate(points[e]):
            probe=point+1e-7*(centers[e]-point);candidates=nearby[e,k]
            candidates=candidates[parent.material_id[candidates]==child.material_id[e]];selected=None
            for ids in [candidates,np.flatnonzero(parent.material_id==child.material_id[e])]:
                bary=np.einsum('nij,j->ni',inverse[ids],np.r_[probe,1.])
                inside=np.flatnonzero(np.min(bary,axis=1)>=-5e-10)
                if len(inside):selected=int(ids[inside[0]]);break
            if selected is None:raise ValueError('Legacy node has no same-material source element')
            a,b,c=inverse[selected]@np.r_[point,1.]
            basis=np.array([a*(2*a-1),b*(2*b-1),c*(2*c-1),4*a*b,4*b*c,4*c*a])
            output[e,k]=basis@source[selected]
    return driver.core.complete_upper_velocity(child,output.ravel(),8)


def gradient_diagnostic(mesh,field,fs):
    area,grad,_=topology(mesh);values=field[:12*len(mesh.tri)].reshape(-1,6,2)
    result={'maximum_gradient_sum':0.,'maximum_raw_minus_anchored_gradient':0.,
        'maximum_raw_minus_wide_gradient':0.,'raw_flow_error':0.,'anchored_flow_error':0.,'wide_flow_error':0.}
    for e,v in enumerate(values):
        c,p=strength(mesh,e,fs);sine=np.sin(p)
        for bary in np.eye(3):
            g=gradients(bary,grad[e]);raw=v.T@g;anchored=(v[1:]-v[0]).T@g[1:]
            wide=np.asarray(v.astype(np.longdouble).T@g.astype(np.longdouble),float)
            result['maximum_gradient_sum']=max(result['maximum_gradient_sum'],float(np.max(abs(g.sum(axis=0)))))
            result['maximum_raw_minus_anchored_gradient']=max(result['maximum_raw_minus_anchored_gradient'],float(np.max(abs(raw-anchored))))
            result['maximum_raw_minus_wide_gradient']=max(result['maximum_raw_minus_wide_gradient'],float(np.max(abs(raw-wide))))
            for name,dv in [('raw',raw),('anchored',anchored),('wide',wide)]:
                error=sine*np.hypot(dv[0,0]-dv[1,1],dv[0,1]+dv[1,0])-np.trace(dv)
                result[name+'_flow_error']=max(result[name+'_flow_error'],float(error))
    return result


def run(case,source,target,output):
    started=time.perf_counter();rec=json.loads((driver.CATALOG/f'{case}.json').read_text(encoding='utf8'))
    parent=driver.load_saved_mesh(source/'best_upper_mesh.npz',rec);child=driver.load_saved_mesh(target/'best_upper_mesh.npz',rec)
    with np.load(source/'best_upper_field.npz',allow_pickle=False) as data:
        field=data['x'].copy();fs=float(data['fs']);assert str(data['mode'])=='upper' and int(data['q'])==8
    ordinary=legacy_global_inverse_interpolation(parent,child,field)
    corrected=interpolate_velocity(parent,child,field,8)
    precise,parents=nested_longdouble_interpolation(parent,child,field)
    prepared=Prepared(child,fs,'upper',8,reduction='none');output.mkdir(parents=True,exist_ok=True)
    records={}
    for label,m,x in [('source',parent,field),('legacy_global_inverse',child,ordinary),
                      ('local_incident',child,corrected),('nested_local',child,precise)]:
        pr=Prepared(m,fs,'upper',8,reduction='none') if label=='source' else prepared
        r=evaluate(pr,x);records[label]={'audit':r,'gradient_diagnostic':gradient_diagnostic(m,x,fs)}
        if label!='source':np.savez_compressed(output/(label+'_field.npz'),x=x,fs=fs,mode='upper',q=8)
    records.update(case=case,fs=fs,elements_source=len(parent.tri),elements_child=len(child.tri),
        maximum_interpolation_change=float(np.max(abs(ordinary-precise))),
        longdouble_eps=float(np.finfo(np.longdouble).eps),
        float64_eps=float(np.finfo(float).eps),
        legacy_selection='first probe candidate with min barycentric >= -5e-10',
        corrected_selection='maximum min barycentric probe toward child centroid',
        full_original_matrix_and_audit_gates_unchanged=True,seconds=time.perf_counter()-started)
    driver.save_mesh(output/'child_mesh.npz',child);write(output/'diagnosis.json',records)
    print(json.dumps({k:{'physical_pass':v['audit']['physical_pass'],'witness':v['audit']['witness'],
        'cone_error':v['audit']['cone_error'],'value':v['audit']['value'],'gradient':v['gradient_diagnostic']}
        for k,v in records.items() if isinstance(v,dict) and 'audit' in v},indent=2),flush=True)
    return records


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--case',required=True)
    parser.add_argument('--source',type=Path,required=True);parser.add_argument('--target',type=Path,required=True)
    parser.add_argument('--output',type=Path,required=True);args=parser.parse_args()
    run(args.case,args.source,args.target,args.output)
