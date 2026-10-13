"""Same-model non-nested flow geometry candidates; no inferred certificate."""
from __future__ import annotations
import math
import experiment_core as core
from experiment_core import np, topology
from model_support.p2 import gradients
from scipy.spatial import cKDTree


def protected_edges(mesh):
    _, _, edges = topology(mesh)
    return {(edge.a, edge.b) for edge in edges if edge.elements[1] < 0 or
            mesh.material_id[edge.elements[0]] != mesh.material_id[edge.elements[1]]}


def audit_geometry(parent, child):
    """Require unchanged coordinates, boundaries, material interfaces and area."""
    if parent.physical_model_id != child.physical_model_id or not np.array_equal(parent.xy, child.xy):
        raise ValueError('Physical model or existing coordinates changed')
    if protected_edges(parent) != protected_edges(child) or set(parent.boundary) != set(child.boundary):
        raise ValueError('Exterior or material interface edges changed')
    for edge, before in parent.boundary.items():
        after=child.boundary[edge]
        if before['kind'] != after['kind'] or not np.array_equal(before['loads'], after['loads']):
            raise ValueError('Boundary condition changed')
    before, _, _ = topology(parent)
    after, _, _ = topology(child)
    pa=np.bincount(parent.material_id,weights=before,minlength=len(parent.materials))
    ca=np.bincount(child.material_id,weights=after,minlength=len(child.materials))
    vertices=child.xy[child.tri]
    ab,ac=vertices[:,1]-vertices[:,0],vertices[:,2]-vertices[:,0]
    if not np.all(ab[:,0]*ac[:,1]-ab[:,1]*ac[:,0]>0) or not np.allclose(pa,ca,rtol=1e-12,atol=1e-10):
        raise ValueError('Orientation or material area changed')
    return {'same_physical_model':True,'protected_edges_preserved':True,
            'coordinates_preserved':True,'material_areas_before':pa.tolist(),
            'material_areas_after':ca.tolist(),'maximum_material_area_difference':float(np.max(abs(pa-ca),initial=0))}


def quality(points, metric=None):
    a,b,c=np.asarray(points)
    ab,ac=b-a,c-a
    determinant=ab[0]*ac[1]-ab[1]*ac[0]
    if determinant <= 0:return -1.
    metric=np.eye(2) if metric is None else metric
    denominator=sum((y-x)@metric@(y-x) for x,y in [(a,b),(b,c),(c,a)])
    return 2*math.sqrt(3)*determinant*math.sqrt(np.linalg.det(metric))/denominator


def oriented(nodes, xy):
    a,b,c=nodes;ab,ac=xy[b]-xy[a],xy[c]-xy[a]
    return [a,b,c] if ab[0]*ac[1]-ab[1]*ac[0]>0 else [a,c,b]


def flow_metric_flip(mesh, record, velocity, passes=5):
    """Archived flow-metric edge flips with unchanged same-material cavities.

    Metrics are sampled on the source velocity and guide element IDs during
    each pass. They select candidates only. The old witness retains its source
    mesh; the candidate requires a new original-program witness and root.
    """
    velocity=np.asarray(velocity,dtype=float)
    if velocity.ndim != 1 or len(velocity)<12*len(mesh.tri) or not np.isfinite(velocity).all():
        raise ValueError('A finite source upper velocity field is required')
    if not isinstance(passes,(int,np.integer)) or passes<0:raise ValueError('Passes must be a nonnegative integer')
    xy,tri=mesh.xy.copy(),mesh.tri.copy()
    _,grad,_=topology(mesh);v=velocity[:12*len(tri)].reshape(-1,6,2)
    metrics=[]
    for e in range(len(tri)):
        jacobian=v[e].T@gradients(np.ones(3)/3,grad[e]);metric=jacobian.T@jacobian
        if not np.isfinite(metric).all():raise ValueError('Nonfinite flow metric')
        val,vec=np.linalg.eigh(metric);high=max(val[-1],1e-20)
        val=np.maximum(val,high/9);metrics.append((vec*val)@vec.T/high)
    changes=[]
    for sweep in range(passes):
        current=core.mesh_from_arrays(xy,tri,mesh.material_id,record)
        _,_,edges=topology(current);candidates=[]
        for edge in edges:
            e,f=edge.elements
            if f<0 or mesh.material_id[e]!=mesh.material_id[f]:continue
            a,b=edge.a,edge.b
            c=next(int(i) for i in tri[e] if i not in [a,b])
            d=next(int(i) for i in tri[f] if i not in [a,b])
            new=[oriented([c,d,a],xy),oriented([d,c,b],xy)]
            def twice_area(cells):
                return sum(abs(np.linalg.det(np.stack([xy[t[1]]-xy[t[0]],xy[t[2]]-xy[t[0]]]))) for t in cells)
            if abs(twice_area([tri[e],tri[f]])-twice_area(new))>1e-10:continue
            metric=(metrics[e]+metrics[f])/2
            before=min(quality(xy[tri[e]],metric),quality(xy[tri[f]],metric))
            after=min(quality(xy[t],metric) for t in new)
            geometric_before=min(quality(xy[tri[e]]),quality(xy[tri[f]]))
            geometric_after=min(quality(xy[t]) for t in new)
            if after>1.05*before+1e-10 and geometric_after>=.6*geometric_before:
                candidates.append((after-before,e,f,new,[a,b],[c,d]))
        if not candidates:break
        used=set()
        for gain,e,f,new,oldedge,newedge in sorted(candidates,reverse=True,key=lambda item:item[0]):
            if e in used or f in used:continue
            tri[e],tri[f]=new;used.update([e,f])
            changes.append({'old':oldedge,'new':newedge,'gain':gain,'sweep':sweep+1})
    child=core.mesh_from_arrays(xy,tri,mesh.material_id,record)
    guard=audit_geometry(mesh,child)
    return child,{'method':'flow_metric_flip','certificate':False,'nested':False,'passes':passes,
                  'flips':len(changes),'elements_before':len(mesh.tri),'elements_after':len(child.tri),
                  'changes':changes,'geometry_audit':guard}


def interpolate_velocity(parent, child, velocity, q=8):
    """Sample a non-nested P2 field within each material, without extrapolation.

    At old jumps the child-centroid direction chooses the incident parent
    triangle. Interpolation alone never implies admissibility. The caller must
    audit all original constraints and complete upper epigraphs at this q.
    Among candidate incident elements, choose the most interior probe. A
    first tolerated match can select the opposite side on very thin elements.
    """
    velocity=np.asarray(velocity,dtype=float)
    if velocity.ndim!=1 or len(velocity)<12*len(parent.tri) or not np.isfinite(velocity).all():
        raise ValueError('A finite source upper velocity field is required')
    if parent.physical_model_id != child.physical_model_id:raise ValueError('Physical model changed')
    source=velocity[:12*len(parent.tri)].reshape(-1,6,2)
    vertices=parent.xy[parent.tri]
    from precision_interpolation import local_barycentric,sample_p2
    centroids=vertices.mean(axis=1);tree=cKDTree(centroids)
    target=child.xy[child.tri]
    points=np.concatenate([target,(target+np.roll(target,-1,axis=1))/2],axis=1)
    target_centroids=target.mean(axis=1)
    _,neighbors=tree.query(points.reshape(-1,2),k=min(32,len(vertices)))
    neighbors=np.asarray(neighbors).reshape(len(child.tri),6,-1)
    output=np.empty((len(child.tri),6,2))
    by_material={int(mid):np.flatnonzero(parent.material_id==mid) for mid in np.unique(parent.material_id)}
    for e in range(len(child.tri)):
        for k,point in enumerate(points[e]):
            candidates=neighbors[e,k]
            candidates=candidates[parent.material_id[candidates]==child.material_id[e]]
            selected=None;best_interior=-np.inf
            for ids in [candidates,by_material[int(child.material_id[e])]]:
                if not len(ids):continue
                at_point=local_barycentric(vertices[ids],np.broadcast_to(point,(len(ids),2)))
                at_center=local_barycentric(vertices[ids],np.broadcast_to(target_centroids[e],(len(ids),2)))
                # The barycentric map is affine. Probe in barycentric space
                # to retain the intended side without adding a tiny offset to
                # large global coordinates.
                bary=at_point+np.longdouble('1e-7')*(at_center-at_point)
                interior=np.min(bary,axis=1);j=int(np.argmax(interior))
                if interior[j]>=-5e-10 and interior[j]>best_interior:
                    selected=int(ids[j]);best_interior=interior[j]
                if best_interior>=0:break
            if selected is None:raise ValueError('Target velocity node lies outside the same-material parent domain')
            output[e,k]=sample_p2(source[selected],local_barycentric(vertices[selected],point))
    return core.complete_upper_velocity(child,output.ravel(),q)
