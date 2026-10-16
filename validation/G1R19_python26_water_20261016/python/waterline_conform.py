"""Conforming geometry for a prescribed, piecewise hydrostatic pressure.

The line is a load discontinuity, not a material interface or velocity glue.
Only geometry is split. Every resulting triangle retains its original soil.
"""
from __future__ import annotations
import numpy as np


def split_at_waterline(xy, tri, material_id, level=3.):
    """Clip each crossing CCW triangle at y=level, sharing edge intersections."""
    xy=np.asarray(xy,dtype=float);tri=np.asarray(tri,dtype=np.int32)
    material_id=np.asarray(material_id,dtype=np.int32)
    if (xy.ndim!=2 or xy.shape[1]!=2 or tri.ndim!=2 or tri.shape[1]!=3 or
            material_id.shape!=(len(tri),) or not np.isfinite(xy).all() or
            not np.isfinite(level)):
        raise ValueError('Finite triangular geometry and material IDs are required')
    points=xy.tolist();intersections={};cells=[];ids=[];parents=[]

    def intersect(a,b):
        if points[a][1]==level:return a
        if points[b][1]==level:return b
        key=tuple(sorted((int(a),int(b))))
        if key not in intersections:
            p=np.asarray(points[key[0]],dtype=np.longdouble)
            q=np.asarray(points[key[1]],dtype=np.longdouble)
            if not min(p[1],q[1])<level<max(p[1],q[1]):
                raise ValueError('A claimed crossing does not straddle the waterline')
            x=p[0]+(np.longdouble(level)-p[1])*(q[0]-p[0])/(q[1]-p[1])
            intersections[key]=len(points);points.append([float(x),float(level)])
        return intersections[key]

    def clipped(nodes,below):
        polygon=[]
        for a,b in zip(nodes,np.roll(nodes,-1)):
            a,b=int(a),int(b)
            ina=points[a][1]<=level if below else points[a][1]>=level
            inb=points[b][1]<=level if below else points[b][1]>=level
            if ina:polygon.append(a)
            if ina!=inb:polygon.append(intersect(a,b))
        unique=[]
        for n in polygon:
            if not unique or unique[-1]!=n:unique.append(n)
        if unique and unique[-1]==unique[0]:unique.pop()
        return unique

    def triangulate(polygon):
        if len(polygon)==3:return [tuple(polygon)]
        if len(polygon)!=4:raise ValueError('A clipped triangle must have3or4vertices')
        a,b,c,d=polygon
        options=[[(a,b,c),(a,c,d)],[(a,b,d),(b,c,d)]]
        def quality(option):
            p=np.asarray(points,dtype=float)[np.asarray(option)]
            u=p[:,1]-p[:,0];v=p[:,2]-p[:,0]
            area=u[:,0]*v[:,1]-u[:,1]*v[:,0]
            edges=np.sum((p-np.roll(p,1,axis=1))**2,axis=(1,2))
            return float(np.min(area/edges))
        return max(options,key=quality)

    crossed=0
    for e,nodes in enumerate(tri):
        y=xy[nodes,1]
        if y.min()<level<y.max():
            crossed+=1
            children=triangulate(clipped(nodes,True))+triangulate(clipped(nodes,False))
        else:children=[tuple(map(int,nodes))]
        cells.extend(children);ids.extend([int(material_id[e])]*len(children));parents.extend([e]*len(children))
    result=np.asarray(points,float),np.asarray(cells,np.int32),np.asarray(ids,np.int32)
    return (*result,{'level':float(level),'elements_before':len(tri),'elements_after':len(cells),
                     'crossing_elements_split':crossed,'new_shared_intersections':len(intersections),
                     'parent_element_ids':parents,'method':'shared_edge_exact_horizontal_waterline_clipping'})


def triangle_moments(xy,tri):
    p=np.asarray(xy)[np.asarray(tri)];a=p[:,1]-p[:,0];b=p[:,2]-p[:,0]
    area=(a[:,0]*b[:,1]-a[:,1]*b[:,0])/2
    return area,area[:,None]*p.mean(axis=1)


def exact_hydrostatic_body(mesh,level=3.,gamma_water=9.81):
    """The same continuum load on every conforming mesh, in physical units."""
    points=mesh.xy[mesh.tri];y=points[:,:,1]
    if np.any((y.min(axis=1)<level)&(y.max(axis=1)>level)):
        raise ValueError('A triangle crosses the waterline; a common exact load is not represented')
    below=y.max(axis=1)<=level
    body=np.zeros_like(mesh.body)
    body[:,1,1]=-mesh.materials[mesh.material_id,3]+gamma_water*below
    # No numerical P1 differentiation is needed: on either side the prescribed
    # pressure is affine and its exact gradient is (0,-gamma_w) or (0,0).
    return body


def verify_partition(parent_xy,parent_tri,parent_ids,child_xy,child_tri,child_ids,change):
    """Check area, first moments, containment and material ownership per parent."""
    pa,pm=triangle_moments(parent_xy,parent_tri);ca,cm=triangle_moments(child_xy,child_tri)
    parent=np.asarray(change['parent_element_ids'],dtype=int)
    if np.any(pa<=0) or np.any(ca<=0):raise ValueError('Nonpositive triangle')
    if not np.array_equal(np.asarray(child_ids),np.asarray(parent_ids)[parent]):
        raise ValueError('Material ownership changed')
    area_error=np.bincount(parent,weights=ca,minlength=len(pa))-pa
    moments=np.column_stack([np.bincount(parent,weights=cm[:,k],minlength=len(pa)) for k in range(2)])
    moment_error=moments-pm
    maximum_outside=0.
    for p,nodes in zip(parent,np.asarray(child_tri)):
        old=np.asarray(parent_xy)[np.asarray(parent_tri)[p]]
        matrix=np.column_stack((old[1]-old[0],old[2]-old[0]))
        bary=np.linalg.solve(matrix,(np.asarray(child_xy)[nodes]-old[0]).T).T
        full=np.column_stack((1-bary.sum(axis=1),bary))
        maximum_outside=max(maximum_outside,float(np.max(-full,initial=0)))
    scale=max(1.,float(np.max(abs(np.asarray(parent_xy)))))
    if np.max(abs(area_error),initial=0)>1e-10*scale**2 or np.max(abs(moment_error),initial=0)>1e-10*scale**3 or maximum_outside>1e-11:
        raise ValueError('Clipped geometry changed parent domain or moments')
    return {'positive_orientation':True,'material_ownership_exact':True,
            'maximum_parent_area_error':float(np.max(abs(area_error),initial=0)),
            'maximum_parent_first_moment_error':float(np.max(abs(moment_error),initial=0)),
            'maximum_barycentric_outside':maximum_outside,'total_area':float(ca.sum())}
