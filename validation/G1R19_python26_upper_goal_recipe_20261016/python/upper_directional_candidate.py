"""Prepare, independently audit and preserve a nested upper candidate.

This experimental CLI makes no solver call and never completes a root. The
one-edge metric algorithm is adapted from the supplied experiment history;
the optional free-vertex fan uses the existing angular enrichment helper.
admissibility is decided only by the original q=8 upper audit.
"""
from pathlib import Path
import argparse
import hashlib
import json

import accuracy_driver as driver
import experiment_core as core
from experiment_core import np
from model_support.p2 import gradients
from precision_interpolation import precision_prolong_upper


def audit_nested_geometry(parent,child):
    if parent.physical_model_id!=child.physical_model_id or not np.array_equal(parent.xy,child.xy[:len(parent.xy)]):
        raise ValueError('Physical model or existing coordinates changed')
    pids=np.asarray(child.provenance['parent_element_ids'],int)
    if not np.array_equal(child.material_id,parent.material_id[pids]) or not np.array_equal(child.body,parent.body[pids]):
        raise ValueError('Child material or body force changed')
    pa,_,pe=core.topology(parent);ca,_,ce=core.topology(child)
    by_parent=np.bincount(pids,weights=ca,minlength=len(pa))
    area_error=float(np.max(abs(by_parent-pa),initial=0.))
    if area_error>1e-9:raise ValueError('Parent triangle area changed')
    old=[edge for edge in pe if edge.elements[1]<0 or parent.material_id[edge.elements[0]]!=parent.material_id[edge.elements[1]]]
    new=[edge for edge in ce if edge.elements[1]<0 or child.material_id[edge.elements[0]]!=child.material_id[edge.elements[1]]]
    starts=np.asarray([parent.xy[edge.a] for edge in old]);directions=np.asarray([parent.xy[edge.b]-parent.xy[edge.a] for edge in old])
    squares=np.sum(directions*directions,axis=1);lengths=np.sqrt(squares);covered=np.zeros(len(old))
    def identity(m,edge):
        e,f=edge.elements
        if f<0:
            data=m.boundary[(edge.a,edge.b)]
            return ('boundary',data['kind'],data['rigid'],tuple(np.asarray(data['loads']).ravel()))
        return ('interface',*sorted((int(m.material_id[e]),int(m.material_id[f]))))
    for edge in new:
        delta=child.xy[[edge.a,edge.b]][:,None,:]-starts[None,:,:]
        position=np.sum(delta*directions[None,:,:],axis=2)/squares[None,:]
        cross=delta[:,:,0]*directions[None,:,1]-delta[:,:,1]*directions[None,:,0]
        matches=np.flatnonzero((np.min(position,axis=0)>=-1e-9)&(np.max(position,axis=0)<=1+1e-9)&
                               (np.max(abs(cross),axis=0)<=1e-9*lengths))
        if len(matches)!=1:raise ValueError('Boundary or interface segment cannot be matched uniquely')
        index=int(matches[0])
        if identity(parent,old[index])!=identity(child,edge):raise ValueError('Boundary load or interface materials changed')
        covered[index]+=edge.length
    length_error=float(np.max(abs(covered-lengths),initial=0.))
    if length_error>1e-9:raise ValueError('Boundary or interface coverage changed')
    return {'pass':True,'existing_coordinates_bitwise_equal':True,'material_and_body_mapping_equal':True,
            'parent_area_error':area_error,'boundary_and_interface_length_error':length_error,
            'source_protected_edges':len(old),'child_protected_edges':len(new),
            'boundary_loads_unchanged':True,'tolerance':1e-9}


def directional_refine(mesh,record,score,velocity,theta=.6):
    if not 0<theta<=1 or len(score)!=len(mesh.tri) or not np.isfinite(score).all() or np.any(score<0):
        raise ValueError('A finite nonnegative indicator and marking fraction are required')
    _,grad,_=core.topology(mesh);values=velocity[:12*len(mesh.tri)].reshape(-1,6,2)
    order=np.argsort(-score);cum=np.cumsum(score[order])
    count=int(np.searchsorted(cum,theta*cum[-1])+1) if cum[-1]>0 else len(order)
    marked=set(map(int,order[:count]));split=set()
    for e in marked:
        J=values[e].T@gradients(np.ones(3)/3,grad[e]);metric=J.T@J
        metric+=np.eye(2)*max(np.trace(metric)*.02,1e-18)
        nodes=mesh.tri[e]
        edges=[mesh.xy[nodes[(k+1)%3]]-mesh.xy[nodes[k]] for k in range(3)]
        k=int(np.argmax([edge@metric@edge for edge in edges]))
        split.add(tuple(sorted([int(nodes[k]),int(nodes[(k+1)%3])])))
    xy=mesh.xy.tolist();mid={}
    for edge in sorted(split):
        mid[edge]=len(xy);xy.append(((mesh.xy[edge[0]]+mesh.xy[edge[1]])/2).tolist())
    tri=[];ids=[];parents=[]
    for e,(t,material) in enumerate(zip(mesh.tri,mesh.material_id)):
        nodes=list(map(int,t));mm=[mid.get(tuple(sorted((nodes[k],nodes[(k+1)%3])))) for k in range(3)]
        number=sum(k is not None for k in mm);a,b,c=nodes
        if number==0:children=[(a,b,c)]
        elif number==3:
            ab,bc,ca=mm;children=[(a,ab,ca),(ab,b,bc),(ca,bc,c),(ab,bc,ca)]
        elif number==1:
            k=next(k for k in range(3) if mm[k] is not None)
            a,b,c=[nodes[(k+j)%3] for j in range(3)];ab=mm[k]
            children=[(a,ab,c),(ab,b,c)]
        else:
            missing=next(k for k in range(3) if mm[k] is None);k=(missing+2)%3
            a,b,c=[nodes[(k+j)%3] for j in range(3)];ab,ca=mm[k],mm[(k+2)%3]
            if np.linalg.norm(np.array(xy[ab])-np.array(xy[c]))<=np.linalg.norm(np.array(xy[b])-np.array(xy[ca])):
                children=[(a,ab,ca),(ab,b,c),(ab,c,ca)]
            else:children=[(a,ab,ca),(ab,b,ca),(b,c,ca)]
        tri.extend(children);ids.extend([int(material)]*len(children));parents.extend([e]*len(children))
    child=core.mesh_from_arrays(np.asarray(xy),np.asarray(tri),np.asarray(ids),record)
    child.provenance['parent_element_ids']=parents
    child,repairs=core.fan_repair(child,record)
    return child,{'method':'directional_one_edge_bisection','marked':len(marked),
                  'split_edges':len(split),'elements_before':len(mesh.tri),'elements_after':len(child.tri),
                  'fan_repairs':repairs,'root_complete':False}


def run(args):
    _,record=driver.load_model(args.case)
    if args.output.exists():raise ValueError('Candidate output must be a new directory')
    parent=driver.load_saved_mesh(args.seed/'best_upper_mesh.npz',record)
    metadata=json.loads((args.seed/'best_upper_certificate.json').read_text(encoding='utf8'))
    with np.load(args.seed/'best_upper_field.npz',allow_pickle=False) as data:
        if str(data['mode'])!='upper' or int(data['q'])!=8:raise ValueError('Only q=8 upper fields are supported')
        x=data['x'].copy();fs=float(data['fs'])
    if (metadata['fs']!=fs or metadata['mesh_id']!=parent.mesh_id or
        metadata['physical_model_id']!=parent.physical_model_id or metadata['case']!=args.case or
        metadata.get('mode')!='upper' or metadata.get('q')!=8):
        raise ValueError('Source field, mesh or model identity mismatch')
    source=core.evaluate(core.Prepared(parent,fs,'upper',8,reduction='none'),x)
    if not source['physical_pass'] or not source['witness']:raise ValueError('Source fails an original physical gate')
    if args.method=='angular_fan':
        from mesh_accuracy import angular_fan_refine
        child,change=angular_fan_refine(parent,record,scope=args.fan_scope)
        change['root_complete']=False
    else:
        child,change=directional_refine(parent,record,core.indicator(parent,x),x,args.theta)
    if len(child.tri)>args.max_elements:raise ValueError('Directional candidate exceeds the element cap')
    change['geometry_audit']=audit_nested_geometry(parent,child)
    prolonged=precision_prolong_upper(parent,child,x,8)
    result=core.evaluate(core.Prepared(child,fs,'upper',8,reduction='none'),prolonged)
    result.update(case=args.case,mode='upper',q=8,fs=fs,mesh_id=child.mesh_id,
                  physical_model_id=child.physical_model_id)
    args.output.mkdir(parents=True)
    driver.save_mesh(args.output/'candidate_mesh.npz',child)
    np.savez_compressed(args.output/'candidate_field.npz',x=prolonged,fs=fs,mode='upper',q=8)
    core.write(args.output/'source_audit.json',source);core.write(args.output/'candidate_audit.json',result)
    change['source_sha256']=dict(driver.SOURCE_HASHES)
    change['source_sha256'].update({str(Path(p).resolve().relative_to(driver.REPO)):hashlib.sha256(Path(p).read_bytes()).hexdigest()
                             for p in [__file__,Path(__file__).with_name('precision_interpolation.py'),
                                       Path(__file__).with_name('mesh_accuracy.py'),
                                       Path(__file__).parent/'vendor/model_support/p2.py']})
    core.write(args.output/'mesh_change.json',change)
    if result['physical_pass'] and result['witness']:
        certificate=driver.Certificate('upper',fs,child,prolonged,result,'audited nested upper candidate',False)
        metadata=certificate.publish(args.output)
        core.write(args.output/'import_audit.json',{'case':args.case,'certificates':{'upper':metadata},'mesh_change':change})
    print('NESTED_UPPER_CANDIDATE',args.case,args.method,len(parent.tri),len(child.tri),
          result['physical_pass'],result['witness'],result['cone_error'],flush=True)


if __name__=='__main__':
    parser=argparse.ArgumentParser();parser.add_argument('--case',required=True)
    parser.add_argument('--seed',type=Path,required=True);parser.add_argument('--output',type=Path,required=True)
    parser.add_argument('--theta',type=float,default=.6);parser.add_argument('--max-elements',type=int,default=24000)
    parser.add_argument('--method',choices=['directional','angular_fan'],default='directional')
    parser.add_argument('--fan-scope',choices=['star','c0'],default='star')
    run(parser.parse_args())
