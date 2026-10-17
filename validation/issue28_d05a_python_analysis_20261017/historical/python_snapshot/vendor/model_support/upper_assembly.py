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

def zero_phi_quadrature(q):
    """Vertex weights on q**2 subtriangles; convex norms give an upper bound.

    The strain of a P2 velocity is affine. Its norm is convex on every
    subtriangle, so the average of its three vertex norms bounds the integral
    from above. This is the existing VBA phi=0 discretization.
    """
    w=np.zeros((q+1,q+1),dtype=int)
    for i in range(q):
        for j in range(q-i):
            w[i,j]+=1;w[i+1,j]+=1;w[i,j+1]+=1
            if i+j<q-1:
                w[i+1,j]+=1;w[i+1,j+1]+=1;w[i,j+1]+=1
    return [(np.array([1-(i+j)/q,i/q,j/q]),w[i,j]/(3*q*q))
            for i in range(q+1) for j in range(q+1-i)]

def complete_upper_velocity(mesh,velocity,q=8):
    """Fill phi=0 epigraphs without changing a prolonged velocity field.

    Auxiliary variables are norms of strain samples and Bernstein jump
    controls. They depend on geometry and velocity, independently of Fs.
    """
    assert 1<=q<=16 and np.all(mesh.rigid_id<0)
    _,grad,edges=topology(mesh);nv=12*len(mesh.tri)
    velocity=np.asarray(velocity,float)[:nv].copy();v=velocity.reshape(-1,6,2)
    auxiliary=[]
    for e in range(len(mesh.tri)):
        if mesh.materials[mesh.material_id[e],1]!=0:continue
        for bary,_ in zero_phi_quadrature(q):
            dv=v[e].T@gradients(bary,grad[e])
            auxiliary.append(math.hypot(dv[0,0]-dv[1,1],dv[0,1]+dv[1,0]))
    for edge in edges:
        e,f=edge.elements
        if f<0 or mesh.material_id[e]!=mesh.material_id[f] or mesh.materials[mesh.material_id[e],1]!=0:continue
        jump=(v[f,edge.loc[1]]-v[e,edge.loc[0]])@edge.tangent
        for segment in range(q):
            auxiliary.extend(abs(bernstein(segment/q,(segment+1)/q)@jump))
    return np.r_[velocity,auxiliary] if auxiliary else velocity

def assemble(mesh,fs,q=2):
    assert 1<=q<=16 and mesh.load_phase=='reference_total'
    assert np.all(mesh.rigid_id<0),'UPPER_RIGID_NOT_IMPLEMENTED'
    area,grad,edges=topology(mesh);n=12*len(mesh.tri);rows=[];dims=[];equal=[]
    work=np.zeros(n);fixed=np.zeros(n);dissipation=np.zeros(n)
    auxiliary_cost=[]
    def auxiliary(price):
        nonlocal n
        i=n;n+=1;auxiliary_cost.append(float(price));return i
    def add_row(ids,values):
        rows.append({int(i):float(v)for i,v in zip(ids,values)if v!=0})
    for e in range(len(mesh.tri)):
        phi=strength(mesh,e,fs);c=reduced_cohesion(mesh,e,fs);ids=np.arange(12*e,12*(e+1))
        for k in range(3):
            g=gradients(np.eye(3)[k],grad[e]);tr=g.ravel();dev=g.copy();dev[:,1]*=-1
            if phi==0:
                equal.append({int(i):float(v) for i,v in zip(ids,tr) if v!=0})
            else:
                dissipation[ids]+=area[e]*c/(3*math.tan(phi))*tr
                add_row(ids,tr/math.sin(phi));add_row(ids,dev.ravel());add_row(ids,g[:,::-1].ravel());dims.append(3)
        if phi==0:
            for bary,weight in zero_phi_quadrature(q):
                g=gradients(bary,grad[e]);dev=g.copy();dev[:,1]*=-1
                rho=auxiliary(area[e]*c*weight)
                add_row([rho],[1.]);add_row(ids,dev.ravel());add_row(ids,g[:,::-1].ravel());dims.append(3)
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
            phi=strength(mesh,e,fs);tangent=math.tan(phi)
            c=reduced_cohesion(mesh,e,fs)
            if phi==0:
                for k in range(3):
                    equal.append({12*edge.elements[side]+2*edge.loc[side][k]+j:
                                  (2*side-1)*edge.normal[j]
                                  for side in range(2) for j in range(2) if edge.normal[j]!=0})
            else:
                for k,weight in enumerate([1,1,4]):
                    for side in range(2):
                        for j in range(2):dissipation[12*edge.elements[side]+2*edge.loc[side][k]+j]+=(2*side-1)*weight*edge.length*c/(6*tangent)*edge.normal[j]
            for segment in range(q):
                weights=bernstein(segment/q,(segment+1)/q)
                for k in range(3):
                    if phi!=0 and not(k==1 or(segment==0 and k==0)or(segment==q-1 and k==2)):continue
                    rho=auxiliary(edge.length*c/(3*q)) if phi==0 else None
                    for sign in [-1,1]:
                        ids=[];values=[];direction=sign*edge.tangent if phi==0 else edge.normal+sign*tangent*edge.tangent
                        for side in range(2):
                            for j in range(3):
                                for component in range(2):
                                    ids.append(12*edge.elements[side]+2*edge.loc[side][j]+component)
                                    values.append((2*side-1)*weights[k,j]*direction[component])
                        if phi==0:ids.append(rho);values.append(1.)
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
    if auxiliary_cost:
        work=np.pad(work,(0,len(auxiliary_cost)));fixed=np.pad(fixed,(0,len(auxiliary_cost)))
        dissipation=np.r_[dissipation,auxiliary_cost]
    def matrix(rr):
        row=[];col=[];data=[]
        for i,r in enumerate(rr):
            for j,v in r.items():row.append(i);col.append(j);data.append(v)
        return sparse.coo_matrix((data,(row,col)),shape=(len(rr),n)).tocsc()
    eq=matrix(equal);norm=np.sqrt(np.asarray(eq.power(2).sum(axis=1)).ravel());assert np.all(norm>0)
    rhs=np.zeros(len(equal));rhs[-1]=1
    return Program(sparse.diags(1/norm)@eq,rhs/norm,matrix(rows),dims,dissipation-fixed,work,norm)
