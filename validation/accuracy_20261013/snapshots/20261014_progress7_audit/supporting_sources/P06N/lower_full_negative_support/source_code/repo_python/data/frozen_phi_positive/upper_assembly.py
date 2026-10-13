from dataclasses import dataclass
import math
import numpy as np
from scipy import sparse
from .schema import topology
from .p2 import gradients,bernstein

def strength(mesh,e,fs):
    c,phi,psi,gamma=mesh.materials[mesh.material_id[e]]
    assert c>=0 and 0<=phi<90 and 0<=psi<=phi and fs>0
    a=math.atan(math.tan(phi*math.pi/180)/fs);b=min(a,max(0,psi*math.pi/180))
    if mesh.provenance.get('psi_policy')=='equal_reduced_phi':b=a
    eta=math.cos(b)*math.cos(a)/(1-math.sin(b)*math.sin(a))
    return math.atan(eta*math.tan(a))

def reduced_cohesion(mesh,e,fs):
    c,phi,psi,gamma=mesh.materials[mesh.material_id[e]]
    a=math.atan(math.tan(phi*math.pi/180)/fs);b=min(a,max(0,psi*math.pi/180))
    if mesh.provenance.get('psi_policy')=='equal_reduced_phi':b=a
    eta=math.cos(b)*math.cos(a)/(1-math.sin(b)*math.sin(a))
    return eta*c/(fs*mesh.stress_scale)

@dataclass
class Program:
    eq:sparse.csc_matrix
    rhs:np.ndarray
    G:sparse.csc_matrix
    dims:list
    objective:np.ndarray
    work:np.ndarray
    row_norm:np.ndarray

def assemble(mesh,fs,q=2):
    assert 1<=q<=16 and mesh.load_phase=='reference_total'
    assert np.all(mesh.rigid_id<0) and np.all(mesh.materials[:,1]>0),'UPPER_PHI_ZERO_NOT_IMPLEMENTED'
    area,grad,edges=topology(mesh);n=12*len(mesh.tri);rows=[];dims=[];equal=[]
    work=np.zeros(n);fixed=np.zeros(n);dissipation=np.zeros(n)
    def add_row(ids,values):
        rows.append({int(i):float(v)for i,v in zip(ids,values)if v!=0})
    for e in range(len(mesh.tri)):
        phi=strength(mesh,e,fs);c=reduced_cohesion(mesh,e,fs);ids=np.arange(12*e,12*(e+1))
        for k in range(3):
            g=gradients(np.eye(3)[k],grad[e]);tr=g.ravel();dev=g.copy();dev[:,1]*=-1
            dissipation[ids]+=area[e]*c/(3*math.tan(phi))*tr
            add_row(ids,tr/math.sin(phi));add_row(ids,dev.ravel());add_row(ids,g[:,::-1].ravel());dims.append(3)
        for k in range(3,6):
            for j in range(2):
                i=12*e+2*k+j;work[i]+=area[e]/3*mesh.body[e,j,1]/mesh.stress_scale;fixed[i]+=area[e]/3*mesh.body[e,j,0]/mesh.stress_scale
    for edge in edges:
        e,f=edge.elements
        if f>=0:
            if mesh.material_id[e]!=mesh.material_id[f]:
                for k in range(3):
                    for j in range(2):equal.append({12*f+2*edge.loc[1][k]+j:1.,12*e+2*edge.loc[0][k]+j:-1.})
                continue
            tangent=math.tan(strength(mesh,e,fs))
            c=reduced_cohesion(mesh,e,fs)
            for k,weight in enumerate([1,1,4]):
                for side in range(2):
                    for j in range(2):dissipation[12*edge.elements[side]+2*edge.loc[side][k]+j]+=(2*side-1)*weight*edge.length*c/(6*tangent)*edge.normal[j]
            for segment in range(q):
                weights=bernstein(segment/q,(segment+1)/q)
                for k in range(3):
                    if not(k==1 or(segment==0 and k==0)or(segment==q-1 and k==2)):continue
                    for sign in [-1,1]:
                        ids=[];values=[];direction=edge.normal+sign*tangent*edge.tangent
                        for side in range(2):
                            for j in range(3):
                                for component in range(2):
                                    ids.append(12*edge.elements[side]+2*edge.loc[side][j]+component)
                                    values.append((2*side-1)*weights[k,j]*direction[component])
                        add_row(ids,values);dims.append(1)
        else:
            assert edge.kind in ['load','fixed','roller_x','roller_y'],'UNSUPPORTED_BOUNDARY'
            for k,weight in enumerate([1,1,4]):
                for j in range(2):
                    i=12*e+2*edge.loc[0][k]+j
                    work[i]+=edge.length*weight/6*edge.loads[j,1]/mesh.stress_scale
                    fixed[i]+=edge.length*weight/6*edge.loads[j,0]/mesh.stress_scale
                    if edge.kind=='fixed'or(edge.kind=='roller_x'and j==0)or(edge.kind=='roller_y'and j==1):equal.append({i:1.})
    equal.append({i:float(v)for i,v in enumerate(work)if abs(v)>=1e-14})
    def matrix(rr):
        row=[];col=[];data=[]
        for i,r in enumerate(rr):
            for j,v in r.items():row.append(i);col.append(j);data.append(v)
        return sparse.coo_matrix((data,(row,col)),shape=(len(rr),n)).tocsc()
    eq=matrix(equal);norm=np.sqrt(np.asarray(eq.power(2).sum(axis=1)).ravel());assert np.all(norm>0)
    rhs=np.zeros(len(equal));rhs[-1]=1
    return Program(sparse.diags(1/norm)@eq,rhs/norm,matrix(rows),dims,dissipation-fixed,work,norm)
