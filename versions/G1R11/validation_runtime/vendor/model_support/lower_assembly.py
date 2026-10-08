"""Soil-only c0 lower operator, including the existing zero-stress reduction."""
import numpy as np
from scipy import sparse
from .schema import topology
from .upper_assembly import strength,reduced_cohesion

def assemble_lower(mesh,fs,fixed_load=None):
    area,grad,edges=topology(mesh);ne=len(mesh.tri);n=18*ne+1;load=n-1
    zero=np.zeros((ne,6),dtype=bool)
    for ed in edges:
        if ed.elements[1]<0 and ed.kind=='load'and np.all(ed.loads==0)and mesh.materials[mesh.material_id[ed.elements[0]],0]==0:zero[ed.elements[0],ed.loc[0]]=True
    changed=True
    while changed:
        changed=False
        for ed in edges:
            e,f=ed.elements
            if f<0 or mesh.materials[mesh.material_id[e],0]!=0 or mesh.materials[mesh.material_id[f],0]!=0:continue
            for a,b in zip(ed.loc[0],ed.loc[1]):
                if zero[e,a]!=zero[f,b]:zero[e,a]=zero[f,b]=True;changed=True
    eq=[];rhs=[];cones=[];offset=[]
    def equal(row,b=0.):
        row={i:v for i,v in row.items()if v!=0}
        if row:eq.append(row);rhs.append(b)
    def add(row,i,value):row[i]=row.get(i,0.)+value
    for e in range(ne):
        for bary in np.eye(3):
            g=np.zeros((6,2))
            for k in range(3):
                j=(k+1)%3;g[k]=2*bary[k]*grad[e,k];g[k+3]=2*(bary[k]*grad[e,j]+bary[j]*grad[e,k])
            for j in range(2):
                row={load:mesh.body[e,j,1]/mesh.stress_scale}
                for k in range(6):add(row,18*e+3*k+j,g[k,j]);add(row,18*e+3*k+2,g[k,1-j])
                equal(row,-mesh.body[e,j,0]/mesh.stress_scale)
        phi=strength(mesh,e,fs);c=reduced_cohesion(mesh,e,fs)
        for k in range(6):
            start=18*e+3*k
            if zero[e,k]:
                for j in range(3):equal({start+j:1.})
            else:
                cones.extend([{start:-np.sin(phi),start+1:-np.sin(phi)},{start:1.,start+1:-1.},{start+2:2.}]);offset.extend([2*c*np.cos(phi),0.,0.])
    def traction(e,k,j,norm):return {18*e+3*k+j:norm[j],18*e+3*k+2:norm[1-j]}
    for ed in edges:
        e,f=ed.elements
        for k in range(3):
            for j in range(2):
                row=traction(e,ed.loc[0][k],j,ed.normal)
                if f>=0:
                    for i,val in traction(f,ed.loc[1][k],j,ed.normal).items():add(row,i,-val)
                    equal(row)
                elif ed.kind=='load'or(ed.kind=='roller_x'and j==1)or(ed.kind=='roller_y'and j==0):
                    add(row,load,-ed.loads[j,1]/mesh.stress_scale);equal(row,ed.loads[j,0]/mesh.stress_scale)
    if fixed_load is not None:equal({load:1.},fixed_load)
    def matrix(rows):
        ii=[];jj=[];vv=[]
        for i,row in enumerate(rows):
            for j,value in row.items():ii.append(i);jj.append(j);vv.append(value)
        return sparse.coo_matrix((vv,(ii,jj)),shape=(len(rows),n)).tocsc()
    A=matrix(eq);norm=np.sqrt(np.asarray(A.power(2).sum(axis=1)).ravel());A=sparse.diags(1/norm)@A;b=np.asarray(rhs)/norm
    return A.tocsc(),b,matrix(cones),np.array(offset),[3]*(len(cones)//3),zero
