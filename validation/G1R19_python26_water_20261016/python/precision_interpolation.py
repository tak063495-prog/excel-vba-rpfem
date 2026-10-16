"""Stable P2 sampling in triangle-local coordinates; physical gates unchanged.

NumPy longdouble is used where available. On Windows it may have the same
precision as float64; the local determinant and partition-of-unity difference
form still avoid cancellation in a global-coordinate 3x3 inverse.
"""
import numpy as np
from scipy.spatial import cKDTree


def local_barycentric(vertices,points):
    vertices=np.asarray(vertices,dtype=np.longdouble)
    points=np.asarray(points,dtype=np.longdouble)
    if vertices.shape[-2:]!=(3,2) or points.shape[-1:]!=(2,):
        raise ValueError('Triangle vertices and sample points have invalid shapes')
    if not np.isfinite(vertices).all() or not np.isfinite(points).all():
        raise ValueError('Interpolation coordinates must be finite')
    a=vertices[...,0,:];u=vertices[...,1,:]-a;v=vertices[...,2,:]-a;w=points-a
    determinant=u[...,0]*v[...,1]-u[...,1]*v[...,0]
    if not np.isfinite(determinant).all() or np.any(determinant==0):
        raise ValueError('Interpolation triangle is degenerate')
    b=(w[...,0]*v[...,1]-w[...,1]*v[...,0])/determinant
    c=(u[...,0]*w[...,1]-u[...,1]*w[...,0])/determinant
    return np.stack([1-b-c,b,c],axis=-1)


def sample_p2(values,bary):
    values=np.asarray(values,dtype=np.longdouble);bary=np.asarray(bary,dtype=np.longdouble)
    if values.shape!=(6,2) or bary.shape[-1:]!=(3,) or not np.isfinite(values).all() or not np.isfinite(bary).all():
        raise ValueError('Finite P2 nodal velocities and barycentric coordinates are required')
    a,b,c=np.moveaxis(bary,-1,0)
    basis=np.stack([b*(2*b-1),c*(2*c-1),4*a*b,4*b*c,4*c*a],axis=-1)
    # Sum of all six P2 basis functions equals one. This form copies a
    # constant velocity exactly and evaluates only local velocity differences.
    return np.asarray(values[0]+basis@(values[1:]-values[0]),dtype=float)


def precision_prolong_upper(parent,child,velocity,q=8,return_parent_ids=False):
    """Optional nested-mesh interpolation, with strict containment checks.

    This returns a candidate field, not a certificate. Every caller must use
    the original physical audit before retaining it as an upper witness.
    """
    from model_support.upper_assembly import complete_upper_velocity
    velocity=np.asarray(velocity,dtype=float)
    if velocity.ndim!=1 or len(velocity)<12*len(parent.tri) or not np.isfinite(velocity).all():
        raise ValueError('A finite source upper field is required')
    if parent.physical_model_id!=child.physical_model_id:
        raise ValueError('Physical model changed during interpolation')
    source=velocity[:12*len(parent.tri)].reshape(-1,6,2)
    vertices=parent.xy[parent.tri];targets=child.xy[child.tri]
    centers=vertices.mean(axis=1);centroids=targets.mean(axis=1)
    points=np.concatenate([targets,(targets+np.roll(targets,-1,axis=1))/2],axis=1)
    provenance=child.provenance.get('parent_element_ids')
    if provenance is not None:
        if len(provenance)!=len(child.tri):raise ValueError('Invalid parent element provenance')
        nearby=np.asarray(provenance,dtype=int).reshape(-1,1)
        if np.any(nearby<0) or np.any(nearby>=len(parent.tri)):raise ValueError('Invalid parent element indices')
    else:
        _,nearby=cKDTree(centers).query(centroids,k=min(32,len(vertices)))
        nearby=np.asarray(nearby).reshape(len(child.tri),-1)
    result=[];parents=[]
    by_material={int(mid):np.flatnonzero(parent.material_id==mid) for mid in np.unique(parent.material_id)}
    for e,center in enumerate(centroids):
        candidates=nearby[e];candidates=candidates[parent.material_id[candidates]==child.material_id[e]]
        selected=None
        for ids in [candidates,by_material.get(int(child.material_id[e]),np.array([],int))]:
            if not len(ids):continue
            bary=local_barycentric(vertices[ids],np.broadcast_to(center,(len(ids),2)))
            for index in np.flatnonzero(np.min(bary,axis=1)>=-1e-12):
                candidate=int(ids[index])
                all_bary=local_barycentric(vertices[candidate],points[e])
                if np.min(all_bary)>=-1e-10:selected=candidate;break
            if selected is not None:break
        if selected is None:raise ValueError('Child is not contained in one same-material source triangle')
        result.append(sample_p2(source[selected],local_barycentric(vertices[selected],points[e])))
        parents.append(selected)
    completed=complete_upper_velocity(child,np.asarray(result).ravel(),q)
    return (completed,parents) if return_parent_ids else completed
