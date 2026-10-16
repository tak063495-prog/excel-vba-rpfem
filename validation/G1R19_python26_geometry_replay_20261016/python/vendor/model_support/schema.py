from dataclasses import dataclass,field
import hashlib
import numpy as np

@dataclass
class Mesh:
    xy:np.ndarray
    tri:np.ndarray
    materials:np.ndarray
    material_id:np.ndarray
    rigid_id:np.ndarray
    body:np.ndarray  # [element, component, fixed/reference]
    boundary:dict   # canonical node pair -> kind, rigid, loads[component,phase]
    stress_scale:float
    physical_model_id:str
    load_phase:str='reference_total'
    provenance:dict=field(default_factory=dict)

    def validate(self):
        n,e=len(self.xy),len(self.tri)
        assert self.xy.shape==(n,2) and self.tri.shape==(e,3)
        assert self.material_id.shape==(e,)and self.rigid_id.shape==(e,)
        assert self.body.shape==(e,2,2)and self.materials.shape[1]==4
        assert np.isfinite(self.xy).all()and np.isfinite(self.body).all()and np.isfinite(self.materials).all()
        assert self.stress_scale>0 and np.min(self.tri)>=0 and np.max(self.tri)<n
        assert np.min(self.material_id)>=0 and np.max(self.material_id)<len(self.materials)
        a=self.xy[self.tri[:,1]]-self.xy[self.tri[:,0]];b=self.xy[self.tri[:,2]]-self.xy[self.tri[:,0]]
        assert np.all(a[:,0]*b[:,1]-a[:,1]*b[:,0]>1e-12),'NONPOSITIVE_TRIANGLE'
        return self

    @property
    def mesh_id(self):
        h=hashlib.sha256()
        for a in [self.xy,self.tri,self.material_id,self.rigid_id,self.body,self.materials]:h.update(np.ascontiguousarray(a).tobytes())
        for key,v in sorted(self.boundary.items()):
            h.update(str(key).encode());h.update(v['kind'].encode());h.update(str(v['rigid']).encode());h.update(np.asarray(v['loads'],float).tobytes())
        h.update(self.load_phase.encode());h.update(np.float64(self.stress_scale).tobytes())
        return h.hexdigest()

@dataclass
class Edge:
    a:int
    b:int
    elements:list
    loc:list
    length:float=0.
    tangent:np.ndarray|None=None
    normal:np.ndarray|None=None
    kind:str='load'
    rigid:int=-1
    loads:np.ndarray=field(default_factory=lambda:np.zeros((2,2)))

def topology(mesh):
    mesh.validate();xy=mesh.xy;tri=mesh.tri
    a=xy[tri[:,1]]-xy[tri[:,0]];b=xy[tri[:,2]]-xy[tri[:,0]];det=a[:,0]*b[:,1]-a[:,1]*b[:,0]
    grad=np.empty((len(tri),3,2))
    for k in range(3):
        a=xy[tri[:,(k+1)%3]];b=xy[tri[:,(k+2)%3]]
        grad[:,k,0]=(a[:,1]-b[:,1])/det;grad[:,k,1]=(b[:,0]-a[:,0])/det
    edges=[];lookup={}
    for e,nodes in enumerate(tri):
        for k in range(3):
            a,b=map(int,[nodes[k],nodes[(k+1)%3]]);key=tuple(sorted((a,b)))
            loc=[k,(k+1)%3,3+k]if a<b else[(k+1)%3,k,3+k]
            if key not in lookup:
                lookup[key]=len(edges);edges.append(Edge(*key,[e,-1],[loc,[-1]*3]))
            else:
                edge=edges[lookup[key]];assert edge.elements[1]==-1,'NONMANIFOLD_EDGE'
                edge.elements[1]=e;edge.loc[1]=loc
    for edge in edges:
        delta=xy[edge.b]-xy[edge.a];edge.length=float(np.linalg.norm(delta));edge.tangent=delta/edge.length
        edge.normal=np.array([edge.tangent[1],-edge.tangent[0]])
        center=xy[tri[edge.elements[0]]].mean(axis=0)-(xy[edge.a]+xy[edge.b])/2
        if center@edge.normal>0:edge.normal*=-1
        if edge.elements[1]<0:
            assert (edge.a,edge.b)in mesh.boundary,'UNCLASSIFIED_EXTERIOR'
            v=mesh.boundary[(edge.a,edge.b)];edge.kind=v['kind'];edge.rigid=v['rigid'];edge.loads=np.array(v['loads'],float)
    return det/2,grad,edges
