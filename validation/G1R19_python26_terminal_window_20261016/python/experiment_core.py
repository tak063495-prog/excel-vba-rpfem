"""Same-model numerical experiments, not a replacement for the Windows solver.

Wall times include Python/native scientific libraries; no VBA speed claims.
All candidate fields are checked on the uncondensed, physical G1R17 mesh.
"""
from __future__ import annotations
import os
for key in ('OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS','NUMEXPR_NUM_THREADS'):
    os.environ[key]='1'
import sys, json, time, math, hashlib, collections, struct
from pathlib import Path
import numpy as np
from scipy import sparse
from scipy.sparse.linalg import lsmr, splu
import clarabel
ROOT=Path(__file__).resolve().parent
sys.path.insert(0,str(ROOT/'vendor'))
from model_support.schema import Mesh, topology
from model_support.lower_assembly import assemble_lower
from model_support.upper_assembly import assemble as assemble_upper, complete_upper_velocity
from independent_audit import audit, strength as audit_strength
from reduction import affine_reduce, affine_reduce_tracked, cone_error

PHYSICAL_TOL=1e-7
RAW_STRESS_TOL=1e-7
LOWER_TARGET=1.000001
UPPER_TARGET=0.999999

def write(path,obj):
    def clean(v):
        if isinstance(v,dict): return {str(k):clean(x) for k,x in v.items()}
        if isinstance(v,(list,tuple)): return [clean(x) for x in v]
        if isinstance(v,np.ndarray): return clean(v.tolist())
        if isinstance(v,np.generic): return clean(v.item())
        if isinstance(v,float) and not math.isfinite(v):return None
        if isinstance(v,Path):return str(v)
        return v
    Path(path).parent.mkdir(parents=True,exist_ok=True)
    Path(path).write_text(json.dumps(clean(obj),ensure_ascii=False,indent=2,allow_nan=False)+'\n')

def mesh_from_arrays(xy,tri,ids,rec):
    mats=np.asarray([[r['c'],r['phi'],r['psi'],r['gamma']] for r in rec['source_case']['materials']],float)
    count=collections.Counter(tuple(sorted((int(a),int(b)))) for t in tri for a,b in zip(t,np.roll(t,-1)))
    kinds={'自由':'load','固定XY':'fixed','固定X':'roller_x','固定Y':'roller_y'};bd={}
    def on(p,a,b):
        a=np.asarray(a);b=np.asarray(b);v=b-a;t=(p-a)@v/(v@v)
        return -1e-10<=t<=1+1e-10 and abs(v[0]*(p[1]-a[1])-v[1]*(p[0]-a[0]))<1e-9*np.linalg.norm(v)
    for edge,num in count.items():
        if num!=1:continue
        matches=[f for f in rec['boundary_explicit'] if all(on(xy[i],f['a'],f['b']) for i in edge)]
        assert len(matches)==1,(edge,matches)
        f=matches[0];loads=np.zeros((2,2));loads[:,1]=np.array(f['load0'])+np.array(f['load1'])
        bd[edge]={'kind':kinds[f['kind']],'rigid':-1,'loads':loads}
    body=np.zeros((len(tri),2,2));body[:,1,1]=-mats[ids,3]
    m=Mesh(np.asarray(xy,float),np.asarray(tri,np.int32),mats,np.asarray(ids,np.int32),
        np.full(len(tri),-1),body,bd,max(1.,float(mats[:,0].max())),
        hashlib.sha256(json.dumps(rec['source_case'],sort_keys=True).encode()).hexdigest(),
        provenance={'psi_policy':rec['source_case']['policy'],'case':rec['case']}).validate()
    area,_,_=topology(m);regions=np.bincount(ids,weights=area,minlength=len(mats))
    assert np.max(np.abs(regions-np.asarray(rec['region_areas'])))<1e-9
    return m

def load_case(case='P06N',original=False):
    rec=json.loads((ROOT/'data'/f'{case}.json').read_text())
    if original:
        cells=json.loads((ROOT/'data/P06N_original243_cells.json').read_text())['cells']
        nc,ec=cells['節点データ'],cells['要素データ']
        nr=sorted(int(k[1:]) for k in nc if k.startswith('A') and k!='A1')
        er=sorted(int(k[1:]) for k in ec if k.startswith('A') and k!='A1')
        xy=np.array([[nc['B'+str(i)],nc['C'+str(i)]] for i in nr],float)
        tri=np.array([[int(ec[c+str(i)])-1 for c in ['B','C','D']] for i in er],np.int32)
        ids=np.array([int(ec['E'+str(i)])-1 for i in er],np.int32)
    else:
        raw=(ROOT/'data/P06N_new_model_mesh.bin').read_bytes();nn,ne=struct.unpack_from('<2i',raw)
        xy=np.frombuffer(raw,dtype='<f8',count=2*nn,offset=8).reshape(nn,2).copy()
        cells=np.frombuffer(raw,dtype='<i4',count=4*ne,offset=8+16*nn).reshape(ne,4)
        tri=cells[:,:3].copy();ids=cells[:,3].copy()
    return mesh_from_arrays(xy,tri,ids,rec),rec

def uniform_refine(m,rec):
    xy=m.xy.tolist();mid={};tri=[];ids=[]
    parent=[]
    for e,(t,mat) in enumerate(zip(m.tri,m.material_id)):
        a,b,c=map(int,t);mm=[]
        for u,v in [(a,b),(b,c),(c,a)]:
            key=tuple(sorted((u,v)))
            if key not in mid:mid[key]=len(xy);xy.append(((m.xy[u]+m.xy[v])/2).tolist())
            mm.append(mid[key])
        ab,bc,ca=mm
        tri.extend([(a,ab,ca),(ab,b,bc),(ca,bc,c),(ab,bc,ca)]);ids.extend([int(mat)]*4);parent.extend([e]*4)
    result=mesh_from_arrays(np.array(xy),np.array(tri),np.array(ids),rec);result.provenance['parent_element_ids']=parent
    return result

def fan_repair(m,rec):
    _,_,_,_,_,zero=assemble_lower(m,1.);candidates=[]
    for e,t in enumerate(m.tri):
        if m.materials[m.material_id[e],0]!=0 or not np.any(m.body[e]):continue
        for k in range(3):
            if not zero[e,k]:continue
            p=m.xy[t[k]];u=m.xy[t[(k+1)%3]]-p;v=m.xy[t[(k+2)%3]]-p
            if u@v < -1e-12*np.linalg.norm(u)*np.linalg.norm(v):
                lu,lv=np.linalg.norm(u),np.linalg.norm(v)
                candidates.append((e,k,p+.5*lv/(lu+lv)*u+.5*lu/(lu+lv)*v));break
    if not candidates:return m,[]
    xy=m.xy.tolist();tri=m.tri.tolist();ids=m.material_id.tolist()
    parent=m.provenance.get('parent_element_ids',list(range(len(tri)))).copy()
    for e,k,p in candidates:
        j=len(xy);xy.append(p.tolist());a,b,c=tri[e];tri[e]=[a,b,j];tri.extend([[b,c,j],[c,a,j]])
        ids.extend([ids[e]]*2);parent.extend([parent[e]]*2)
    result=mesh_from_arrays(np.array(xy),np.array(tri),np.array(ids),rec);result.provenance['parent_element_ids']=parent
    return result,[{'element':e,'vertex':k,'point':p.tolist()} for e,k,p in candidates]

def indicator(m,x):
    from model_support.p2 import gradients
    area,grad,edges=topology(m);v=np.asarray(x)[:12*len(m.tri)].reshape(-1,6,2);score=np.zeros(len(m.tri))
    for e in range(len(m.tri)):
        for k in range(3):
            dv=v[e].T@gradients(np.eye(3)[k],grad[e]);score[e]+=area[e]*math.hypot(dv[0,0]-dv[1,1],dv[0,1]+dv[1,0])/3
    for edge in edges:
        e,f=edge.elements
        if f<0:continue
        jump=v[f,edge.loc[1]]-v[e,edge.loc[0]]
        val=edge.length*np.array([1,1,4])@np.linalg.norm(jump,axis=1)/6
        score[e]+=val/2;score[f]+=val/2
    return score

def local_refine(m,rec,score,theta=.6):
    order=np.argsort(-score);cum=np.cumsum(score[order]);count=int(np.searchsorted(cum,theta*cum[-1])+1) if cum[-1]>0 else len(order)
    marked=set(map(int,order[:count]));split=set()
    for e in marked:
        a,b,c=map(int,m.tri[e]);split.update(tuple(sorted(k)) for k in [(a,b),(b,c),(c,a)])
    xy=m.xy.tolist();mid={}
    for edge in sorted(split):mid[edge]=len(xy);xy.append(((m.xy[edge[0]]+m.xy[edge[1]])/2).tolist())
    tri=[];ids=[];parent=[]
    for e,(t,mat) in enumerate(zip(m.tri,m.material_id)):
        a,b,c=map(int,t);nodes=[a,b,c];mm=[mid.get(tuple(sorted((nodes[k],nodes[(k+1)%3])))) for k in range(3)];ns=sum(v is not None for v in mm)
        if ns==0:children=[(a,b,c)]
        elif ns==3:
            ab,bc,ca=mm;children=[(a,ab,ca),(ab,b,bc),(ca,bc,c),(ab,bc,ca)]
        elif ns==1:
            k=next(j for j in range(3) if mm[j] is not None);a,b,c=nodes[k],nodes[(k+1)%3],nodes[(k+2)%3];ab=mm[k]
            children=[(a,ab,c),(ab,b,c)]
        else:
            missing=next(j for j in range(3) if mm[j] is None);k=(missing+2)%3
            a,b,c=nodes[k],nodes[(k+1)%3],nodes[(k+2)%3];ab,ca=mm[k],mm[(k+2)%3]
            # Choose the shorter diagonal in the remaining quadrilateral.
            if np.linalg.norm(np.array(xy[ab])-np.array(xy[c]))<=np.linalg.norm(np.array(xy[b])-np.array(xy[ca])):
                children=[(a,ab,ca),(ab,b,c),(ab,c,ca)]
            else:children=[(a,ab,ca),(ab,b,ca),(b,c,ca)]
        tri.extend(children);ids.extend([int(mat)]*len(children));parent.extend([e]*len(children))
    new=mesh_from_arrays(np.array(xy),np.array(tri),np.array(ids),rec);new.provenance['parent_element_ids']=parent
    new,repairs=fan_repair(new,rec)
    return new,{'marked':len(marked),'theta':theta,'elements_before':len(m.tri),'elements_after':len(new.tri),'fan_repairs':repairs}

def prolong(parent,child,x,mode,q=8):
    result=[];pids=child.provenance['parent_element_ids'];values=x[:18*len(parent.tri)].reshape(-1,6,3) if mode=='lower' else x[:12*len(parent.tri)].reshape(-1,6,2)
    for t,e in zip(child.tri,pids):
        vertices=child.xy[t];points=np.r_[vertices,(vertices+np.roll(vertices,-1,axis=0))/2]
        matrix=np.vstack([parent.xy[parent.tri[e]].T,np.ones(3)])
        bary=np.linalg.solve(matrix,np.vstack([points.T,np.ones(6)])).T
        a,b,c=bary.T
        if mode=='lower':
            B=np.column_stack([a*a,b*b,c*c,2*a*b,2*b*c,2*c*a]);v=B@values[e]
            v[3:]=2*v[3:]-.5*(v[:3]+np.roll(v[:3],-1,axis=0));result.append(v)
        else:
            N=np.column_stack([a*(2*a-1),b*(2*b-1),c*(2*c-1),4*a*b,4*b*c,4*c*a]);result.append(N@values[e])
    return np.r_[np.array(result).ravel(),x[-1]] if mode=='lower' else complete_upper_velocity(child,np.array(result).ravel(),q)

def program(m,fs,mode,q=8,fixed_load=None,include_row_norm=False):
    if mode=='lower':
        assembled=assemble_lower(m,fs,fixed_load)
        E,b,G,h,dims,zero=assembled
        row_norm=assembled.row_norm
        cost=np.zeros(E.shape[1]);cost[-1]=-1. if fixed_load is None else 0.
    else:
        p=assemble_upper(m,fs,q);E,b,G,h,dims,cost=p.eq,p.rhs,p.G,np.zeros(p.G.shape[0]),p.dims,p.objective
        row_norm=p.row_norm
    result=E.tocsc(),np.asarray(b),G.tocsc(),np.asarray(h),list(dims),cost
    return (result,row_norm) if include_row_norm else result

class Prepared:
    def __init__(self,m,fs,mode,q=8,reduction='none',fixed_load=None):
        self.m,self.fs,self.mode,self.q,self.kind=m,fs,mode,q,reduction
        t=time.perf_counter()
        if mode=='lower':
            self.original,self.original_row_norm=program(m,fs,mode,q,fixed_load,include_row_norm=True)
        else:
            self.original=program(m,fs,mode,q,fixed_load)
            self.original_row_norm=None
        self.assembly_seconds=time.perf_counter()-t
        E,b,G,h,dims,cost=self.original;self.T=None;self.off=np.zeros(len(cost));self.cone_ids=np.arange(len(dims));self.stats=[]
        # Bind this physical optimization objective and layout before any
        # auxiliary feasibility program can copy and change this preparation.
        self.physical_original_objective=cost.copy()
        self.physical_original_objective.setflags(write=False)
        self.physical_original_layout=(E.shape,G.shape,len(b),len(h),tuple(dims))
        t=time.perf_counter()
        if reduction=='zero':
            er=E.tocsr();single=np.flatnonzero(np.diff(er.indptr)==1)
            cols=np.unique([er.indices[er.indptr[i]] for i in single if b[i]==0])
            keep=np.setdiff1d(np.arange(len(cost)),cols)
            self.T=sparse.eye(len(cost),format='csc')[:,keep];EE=(E@self.T).tocsc();GG=(G@self.T).tocsc()
            active=np.asarray(EE.getnnz(axis=1)).ravel()!=0
            assert np.max(np.abs(b[~active]),initial=0)<2e-11
            E,b,G=EE[active],b[active],GG
            # Keep constant cones for a simple exact map; Clarabel may presolve them.
            self.stats=[{'removed_variables':len(cols),'removed_equalities':int((~active).sum())}]
        elif reduction=='affine':
            if mode=='lower':
                E,b,G,h,dims,self.T,self.off,self.stats,self.history,self.final_keep=affine_reduce_tracked(E,b,G,h,dims,cost)
            else:E,b,G,h,dims,self.T,self.off,self.stats=affine_reduce(E,b,G,h,dims)
            self.cone_ids=np.asarray(self.stats[-1]['retained_original_cones'],int)
        elif reduction!='none':raise ValueError(reduction)
        self.cost=np.asarray(cost@self.T).ravel() if self.T is not None else cost.copy()
        self.objective_offset=float(cost@self.off)
        self.reduced=E.tocsc(),b,G.tocsc(),h,dims,self.cost
        self.transform_seconds=time.perf_counter()-t
        self.A=sparse.vstack([E,-G],format='csc');self.rhs=np.r_[b,h]
        self.cones=[clarabel.ZeroConeT(len(b))]+[clarabel.NonnegativeConeT(d) if d==1 else clarabel.SecondOrderConeT(d) for d in dims]
    def restore(self,x):
        return np.asarray(self.T@x+self.off).ravel() if self.T is not None else np.asarray(x).copy()
    def restore_dual(self,z):
        E0,b0,G0,h0,dims0,cost0=self.original;E,b,G,h,dims,cost=self.reduced
        if self.T is None:return np.asarray(z),{'method':'identity','seconds':0.}
        if hasattr(self,'history'):
            t=time.perf_counter();yy=np.zeros(self.history[-1]['E'].shape[0]);yy[self.final_keep]=z[:len(b)];zz=z[len(b):].copy()
            zorig=np.zeros(G0.shape[0]);orig_off=np.r_[0,np.cumsum(dims0)];k=len(b)
            for j,d in zip(self.cone_ids,dims):zorig[orig_off[j]:orig_off[j+1]]=z[k:k+d];k+=d
            fullgradient=cost0-G0.T@zorig
            for step_index in reversed(range(len(self.history))):
                step=self.history[step_index]
                EP,GP=step['E'],step['G'];zc=np.zeros(GP.shape[0]);zc[step['cone_keep']]=zz
                residual=np.asarray(step['T'].T@fullgradient).ravel()+EP.T@yy;ec=EP.tocsr();adj={};singles=[]
                for row in step['pivots']:
                    lo,hi=ec.indptr[row:row+2];ii=ec.indices[lo:hi];vv=ec.data[lo:hi]
                    if len(ii)==1:singles.append((int(row),int(ii[0]),float(vv[0])))
                    elif len(ii)==2:
                        i,j=map(int,ii);a,c=map(float,vv)
                        adj.setdefault(i,[]).append((j,int(row),a,c));adj.setdefault(j,[]).append((i,int(row),c,a))
                    else:raise ValueError('NON_GRAPH_PIVOT')
                for row,i,v in singles:
                    dy=-residual[i]/v;yy[row]+=dy;residual[i]+=v*dy
                roots=set(map(int,step['roots']));degree={i:len(a) for i,a in adj.items()};used=set()
                queue=collections.deque(i for i,d in degree.items() if d==1 and i not in roots)
                while queue:
                    i=queue.popleft()
                    if degree[i]!=1:continue
                    j,row,a,c=next(edge for edge in adj[i] if edge[1] not in used)
                    dy=-residual[i]/a;yy[row]+=dy;residual[i]+=a*dy;residual[j]+=c*dy
                    used.add(row);degree[i]-=1;degree[j]-=1
                    if degree[j]==1 and j not in roots:queue.append(j)
                assert len(used)==sum(len(a) for a in adj.values())//2
                previous_size=self.original[0].shape[0] if step_index==0 else self.history[step_index-1]['E'].shape[0]
                # input_keep indexes the equality rows entering this sweep.
                prev=np.zeros(previous_size);prev[step['input_keep']]=yy;yy=prev;zz=zc
            z0=np.r_[yy,zz]
            return z0,{'method':'reverse_elimination_graph','seconds':time.perf_counter()-t,'sweeps':len(self.history)}
        t=time.perf_counter();z0=np.zeros(len(b0)+len(h0));offs0=np.r_[0,np.cumsum(dims0)];k=len(b)
        for j,d in zip(self.cone_ids,dims):
            z0[len(b0)+offs0[j]:len(b0)+offs0[j+1]]=z[k:k+d];k+=d
        EE=(E0@self.T).tocsr();EE.eliminate_zeros();active=np.flatnonzero(np.diff(EE.indptr)>0)
        compatible=len(active)==len(b) and EE[active].shape==E.shape
        if compatible:
            difference=EE[active]-E;compatible=np.max(np.abs(difference.data),initial=0)<2e-10
        if compatible:
            z0[active]=z[:len(b)];removed=np.setdiff1d(np.arange(len(b0)),active)
            residual=cost0+E0.T@z0[:len(b0)]-G0.T@z0[len(b0):]
            M=E0[removed].T.tocsc();rhs=-residual
            ls=lsmr(M,rhs,atol=1e-13,btol=1e-13,maxiter=5000)
            z0[removed]=ls[0];method='removed_equalities_lsmr'
        else:
            # Face-exposed conic duals need additional restoration; do not claim
            # validity if this least-squares attempt leaves a stationarity defect.
            rhs=-cost0+G0.T@z0[len(b0):]
            ls=lsmr(E0.T,rhs,atol=1e-13,btol=1e-13,maxiter=10000)
            z0[:len(b0)]=ls[0];method='full_equalities_lsmr'
        return z0,{'method':method,'seconds':time.perf_counter()-t,'lsmr_flag':ls[1],'lsmr_iterations':ls[2]}

def evaluate(pr,x,z=None):
    E,b,G,h,dims,cost=pr.original;m=pr.m
    if not np.isfinite(x).all():return {'physical_pass':False,'numerical_gate':False,'witness':False,
        'value':float('nan'),'objective':float('nan'),'eq_error':float('inf'),'cone_error':float('inf'),
        'audit':{'pass':False,'errors':{}},'raw_lower_errors':{},'reason':'NONFINITE'}
    a=audit(m,x,pr.mode,pr.fs,pr.q,float(cost@x),18*len(m.tri))
    eq=float(np.max(np.abs(E@x-b),initial=0));cone=cone_error(G@x+h,dims)
    raw={k:float(v*m.stress_scale) for k,v in a['errors'].items()} if pr.mode=='lower' else {}
    physical=bool(a['pass'] and eq<=PHYSICAL_TOL and cone<=PHYSICAL_TOL and max(raw.values(),default=0)<=RAW_STRESS_TOL)
    r={'objective':float(cost@x),'value':float(a['value']),'eq_error':eq,'cone_error':cone,
       'audit':a,'raw_lower_errors':raw,'physical_pass':physical,
       'witness':bool(physical and (a['value']>=LOWER_TARGET if pr.mode=='lower' else a['value']<=UPPER_TARGET))}
    if z is not None and np.isfinite(z).all():
        y=z[:len(b)];zz=z[len(b):];dual=-float(b@y+h@zz)
        stat=float(np.max(np.abs(cost+E.T@y-G.T@zz),initial=0));dc=cone_error(zz,dims)
        gap=abs(float(cost@x)-dual)
        r.update(dual_objective=dual,dual_stationarity=stat,dual_cone_error=dc,dual_gap=gap,
                 numerical_gate=bool(physical and eq<1e-8 and stat<1e-7 and dc<=1e-8 and gap<1e-8*(1+abs(float(cost@x)))))
    else:r['numerical_gate']=False
    return r

def settings(overrides=None,max_iter=250):
    s=clarabel.DefaultSettings();s.verbose=False;s.max_threads=1;s.max_iter=max_iter;s.time_limit=45.
    s.tol_gap_abs=s.tol_gap_rel=s.tol_feas=s.tol_infeas_abs=s.tol_infeas_rel=1e-10
    for k,v in (overrides or {}).items():setattr(s,k,v)
    return s

def inverse_dual_guide(pr,result):
    """Verify an original optimization dual that excludes a primal bound.

    This guide screens a discrete root only. Fixed-load lower and modified
    upper feasibility objectives cannot suppress physical witness recovery.
    """
    E,b,G,h,dims,cost=pr.original
    values=[result.get(key) for key in ('dual_stationarity','dual_cone_error','dual_objective')]
    try:finite=bool(np.isfinite(np.asarray(values,dtype=float)).all())
    except (TypeError,ValueError):finite=False
    evidence={'verified':False,'scope':'discrete_inverse_sign_only; primal unchanged'}
    if not finite:return dict(evidence,reason='NONFINITE_OR_MISSING_ORIGINAL_DUAL')
    stationarity,cone,objective=map(float,values)
    evidence.update(dual_stationarity=stationarity,dual_cone_error=cone,dual_objective=objective)
    if stationarity<0 or cone<0 or stationarity>1e-7 or cone>1e-8:
        return dict(evidence,reason='ORIGINAL_RAW_DUAL_GATE_REJECTED')
    if pr.mode=='lower':
        expected=np.zeros(18*len(pr.m.tri)+1);expected[-1]=-1.
        if not np.array_equal(cost,expected):return dict(evidence,reason='LOWER_FREE_LOAD_OBJECTIVE_AND_LAYOUT_REQUIRED')
        verified=-objective<=LOWER_TARGET-2e-6
    elif pr.mode=='upper':
        layout=(E.shape,G.shape,len(b),len(h),tuple(dims))
        if layout!=getattr(pr,'physical_original_layout',None) or not np.array_equal(cost,getattr(pr,'physical_original_objective',None)):
            return dict(evidence,reason='ORIGINAL_UPPER_OBJECTIVE_AND_LAYOUT_REQUIRED')
        verified=objective>=UPPER_TARGET+2e-6
    else:return dict(evidence,reason='LOWER_OR_UPPER_ONLY')
    evidence.update(verified=bool(verified),reason='STRICT_ORIGINAL_INVERSE_DUAL_GUIDE' if verified else 'NO_INVERSE_SEPARATION')
    return evidence

def cone_row_scales(matrix,dimensions):
    """One positive Frobenius multiplier per original cone block.

    Row slicing a CSC matrix once per small cone scans every column. Sum the
    squares in CSR once instead, then aggregate the contiguous cone rows.
    This changes neither the cone nor the original-program audit.
    """
    dimensions=np.asarray(dimensions,dtype=np.int64)
    if dimensions.ndim!=1 or np.any(dimensions<=0) or int(dimensions.sum())!=matrix.shape[0]:
        raise ValueError('Positive cone dimensions must cover the matrix rows')
    if len(dimensions)==0:return np.ones(matrix.shape[0])
    rows=matrix.tocsr()
    squares=np.asarray(rows.multiply(rows).sum(axis=1)).ravel()
    starts=np.r_[0,np.cumsum(dimensions[:-1])]
    norms=np.sqrt(np.add.reduceat(squares,starts))
    multipliers=np.ones(len(dimensions))
    active=norms>1e-10
    multipliers[active]=1/norms[active]
    return np.repeat(multipliers,dimensions)


def polish_primal(pr,source,dual=None,before=None):
    """Try lower physical row weights, then retain the original recovery path.

    A positive diagonal scaling changes only the correction search. Every
    candidate must pass the uncondensed program and the independent original
    physics audit. Failed corrections never replace or mutate the source.
    """
    source=np.asarray(source,dtype=float)
    E,b,*_=pr.original
    if source.shape!=(E.shape[1],) or not np.isfinite(source).all():
        raise ValueError('Primal recovery requires a finite original-program field')
    before=evaluate(pr,source,dual) if before is None else before
    if before['physical_pass']:
        return {'accepted':False,'skipped':True,'reason':'SOURCE_ALREADY_PHYSICAL','seconds':0.},before,source
    started=time.perf_counter()
    weighted={'attempted':False,'accepted':False,'stages':[]}
    row_norm=getattr(pr,'original_row_norm',None)
    if pr.mode=='lower' and row_norm is not None:
        row_norm=np.asarray(row_norm,dtype=float)
        if row_norm.shape==(len(b),) and np.isfinite(row_norm).all() and np.all(row_norm>0):
            matrix=(sparse.diags(row_norm)@E).tocsc()
            rhs=row_norm*(b-E@source)
            if np.isfinite(matrix.data).all() and np.isfinite(rhs).all():
                weighted.update(attempted=True,method='original_lower_physical_equality_row_weights',
                    minimum_row_weight=float(np.min(row_norm)),
                    maximum_row_weight=float(np.max(row_norm)))
                correction=np.zeros(E.shape[1]);iterations=0
                for budget in [100,500,1000]:
                    stage_budget=budget-iterations
                    tick=time.perf_counter()
                    step=lsmr(matrix,rhs,atol=1e-15,btol=1e-15,maxiter=stage_budget,x0=correction)
                    correction=step[0];iterations+=step[2]
                    candidate=source+correction
                    result=evaluate(pr,candidate,dual)
                    weighted['stages'].append({'cumulative_budget':budget,'iterations':iterations,
                        'lsmr_flag':step[1],'seconds':time.perf_counter()-tick,
                        'physical_pass':result['physical_pass'],'witness':result['witness'],
                        'value':result['value'],'eq_error':result['eq_error'],'cone_error':result['cone_error'],
                        'raw_lower_errors':result['raw_lower_errors'],
                        'correction_norm':float(np.linalg.norm(correction))})
                    if result['physical_pass']:
                        weighted['accepted']=True
                        return {'method':'lower_physical_row_lsmr_then_full_original_audit','accepted':True,
                            'lsmr_flag':step[1],'iterations':iterations,'correction_norm':float(np.linalg.norm(correction)),
                            'seconds':time.perf_counter()-started,'before_raw':before.get('raw_lower_errors'),
                            'after_raw':result.get('raw_lower_errors'),'weighted_lower_recovery':weighted},result,candidate
                    if step[1] in (0,1,2,4,5) and step[2]<stage_budget:
                        weighted['converged_without_physical_certificate']=True
                        break
            else:weighted['reason']='NONFINITE_WEIGHTED_CORRECTION_PROGRAM'
        else:weighted['reason']='INVALID_ORIGINAL_ROW_NORMS'
    # The previous unweighted recovery starts from the unchanged source,
    # retaining its full original 10000-iteration budget on every failure.
    correction=lsmr(E,b-E@source,atol=1e-15,btol=1e-15,maxiter=10000)
    candidate=source+correction[0];result=evaluate(pr,candidate,dual)
    accepted=bool(result['physical_pass'])
    record={'method':'original_equality_lsmr_then_full_physical_audit','accepted':accepted,
        'lsmr_flag':correction[1],'iterations':correction[2],
        'correction_norm':float(np.linalg.norm(correction[0])),
        'seconds':time.perf_counter()-started,'before_raw':before.get('raw_lower_errors'),
        'after_raw':result.get('raw_lower_errors')}
    if pr.mode=='lower':record.update(weighted_lower_recovery=weighted,fallback_from_unchanged_source=True)
    return record,result if accepted else before,candidate if accepted else source


def solve(pr,overrides=None,max_iter=250,duals=True,save=None,early=False,early_primal_limit=1e-5,cone_scaling=False,primal_polish=False):
    row_scale=np.ones(pr.A.shape[0])
    if cone_scaling:
        Er,br,Gr,hr,dr,cr=pr.reduced
        row_scale[len(br):]=cone_row_scales(Gr,dr)
    AA=(sparse.diags(row_scale)@pr.A).tocsc() if cone_scaling else pr.A
    bb=row_scale*pr.rhs if cone_scaling else pr.rhs
    t=time.perf_counter();s=clarabel.DefaultSolver(sparse.csc_matrix((len(pr.cost),len(pr.cost))),pr.cost,AA,bb,pr.cones,settings(overrides,max_iter))
    build=time.perf_counter()-t
    trigger=[]
    if early:
        def callback(info):
            value=-(info.cost_primal+pr.objective_offset) if pr.mode=='lower' else info.cost_primal+pr.objective_offset
            meaning=value>=LOWER_TARGET+2e-6 if pr.mode=='lower' else value<=UPPER_TARGET-2e-6
            stop=bool(meaning and info.res_primal<=early_primal_limit)
            if stop:trigger.append({'iteration':info.iterations,'value':value,'relative_primal_residual':info.res_primal})
            return stop
        s.set_termination_callback(callback)
    t=time.perf_counter();sol=s.solve();wall=time.perf_counter()-t
    x=pr.restore(np.asarray(sol.x));z=None;rest={};canonical_z=row_scale*np.asarray(sol.z)
    if duals:z,rest=pr.restore_dual(canonical_z)
    t=time.perf_counter();r=evaluate(pr,x,z);au_time=time.perf_counter()-t
    polish={}
    if primal_polish and not r['physical_pass'] and np.isfinite(x).all():
        guide=inverse_dual_guide(pr,r)
        if guide['verified']:
            polish={'method':'original_equality_lsmr_then_full_physical_audit','skipped':True,
                    'accepted':False,'reason':'STRICT_ORIGINAL_INVERSE_DUAL_GUIDE','raw_original_dual_evidence':guide,
                    'primal_unchanged':True,'seconds':0.}
        else:
            polish,r,x=polish_primal(pr,x,z,r)
    # For a strictly zero-cost upper program, a feasible Farkas ray can be
    # multiplied by a positive scalar without changing its original-program
    # meaning. Clarabel commonly normalizes its separation to one, just below
    # the driver's positive root margin. Re-audit the doubled, restored ray in
    # the uncondensed program; it supplies a discrete root guide only.
    ray_rescaling={}
    if (duals and pr.mode=='upper' and
            str(sol.status) in ('PrimalInfeasible','AlmostPrimalInfeasible') and
            np.all(pr.original[-1]==0) and z is not None and
            np.isfinite(z).all() and np.isfinite(canonical_z).all()):
        E0,b0,G0,h0,dims0,cost0=pr.original
        def ray_metrics(original_z):
            y0,zc0=original_z[:len(b0)],original_z[len(b0):]
            return {'objective':-float(b0@y0+h0@zc0),
                    'stationarity':float(np.max(np.abs(cost0+E0.T@y0-G0.T@zc0),initial=0)),
                    'cone_error':cone_error(zc0,dims0)}
        with np.errstate(over='ignore',invalid='ignore'):
            before=ray_metrics(z);candidate_z=2*z;candidate_canonical=2*canonical_z
            after=ray_metrics(candidate_z) if (np.isfinite(candidate_z).all() and
                np.isfinite(candidate_canonical).all()) else None
        accepted=bool(after is not None and all(math.isfinite(v) for v in before.values()) and
            all(math.isfinite(v) for v in after.values()) and
            after['stationarity']<=1e-7 and after['cone_error']<=1e-8 and
            after['objective']>before['objective'])
        ray_rescaling={'factor':2.,'accepted':accepted,'before':before,'after':after,
            'scope':'DISCRETE_UPPER_INFEASIBILITY_GUIDE_ONLY'}
        if accepted:
            z,canonical_z=candidate_z,candidate_canonical
            r.update(dual_objective=after['objective'],dual_stationarity=after['stationarity'],
                     dual_cone_error=after['cone_error'],
                     dual_gap=abs(float(cost0@x)-after['objective']),numerical_gate=False)
    # Reduced dual metrics are diagnostic. An elimination can have a valid
    # reduced dual even when restoration leaves an original-program defect;
    # only the fully restored and audited original dual can guide root search.
    # A published continuum bound still requires an audited primal field.
    Er,br,Gr,hr,dr,cr=pr.reduced;zr=canonical_z;yr=zr[:len(br)];zc=zr[len(br):]
    reduced_dual=-float(br@yr+hr@zc)+pr.objective_offset
    reduced_stationarity=float(np.max(np.abs(cr+Er.T@yr-Gr.T@zc),initial=0))
    reduced_cone=cone_error(zc,dr);reduced_gap=abs(float(pr.original[-1]@x)-reduced_dual)
    r.update(reduced_dual_objective=reduced_dual,reduced_dual_stationarity=reduced_stationarity,
             reduced_dual_cone_error=reduced_cone,reduced_dual_gap=reduced_gap,
             reduced_numerical_gate=bool(r['physical_pass'] and reduced_stationarity<=1e-7 and
               reduced_cone<=1e-8 and reduced_gap<=1e-8*(1+abs(float(pr.original[-1]@x)))))
    r.update(case=pr.m.provenance['case'],elements=len(pr.m.tri),fs=pr.fs,mode=pr.mode,q=pr.q,reduction=pr.kind,
             status=str(sol.status),iterations=sol.iterations,solver_seconds=wall,build_seconds=build,
             backend_solve_seconds=sol.solve_time,audit_seconds=au_time,dual_restore=rest,
             assembly_seconds=pr.assembly_seconds,transform_seconds=pr.transform_seconds,
             original_variables=pr.original[0].shape[1],original_equalities=pr.original[0].shape[0],
             variables=len(pr.cost),equalities=len(pr.reduced[1]),cones=len(pr.reduced[4]),
             solver_reported_gap=abs(sol.obj_val-sol.obj_val_dual),settings_overrides=overrides or {},max_iter=max_iter,
             early=early,callback_trigger=trigger,cone_scaling=cone_scaling,primal_polish=polish)
    if ray_rescaling:r['original_dual_ray_rescaling']=ray_rescaling
    r['end_to_end_seconds']=pr.assembly_seconds+pr.transform_seconds+build+wall+au_time+rest.get('seconds',0)+polish.get('seconds',0)
    if save:
        np.savez_compressed(save,x=x,z=z if z is not None else np.array([]),xr=np.asarray(sol.x),zr=canonical_z,fs=pr.fs,mode=pr.mode,q=pr.q)
    return r,x,canonical_z

class LowerRefresh:
    """Reuse lower E and reduction maps, refreshing the true Fs-dependent cones.
    The existing G1R17 already prepares its structure; this measures the extra
    reduction map once per mesh, not a new benefit attributed to assembly reuse.
    """
    def __init__(self,m,reduction='none',fs=.5,fixed_load=None):
        self.base=Prepared(m,fs,'lower',reduction=reduction,fixed_load=fixed_load);self.m=m
        E,b,G,h,dims,cost=self.base.original
        assert all(d==3 for d in dims)
        assert not any(s.get('zero_faces',0) for s in self.base.stats)
        self.head_data_mask=G.indices%3==0
        cols=np.repeat(np.arange(G.shape[1]),np.diff(G.indptr))
        self.head_data_owners=cols[self.head_data_mask]//18
        gc=G.tocsr();self.cone_owners=np.array([gc.indices[gc.indptr[k]]//18 for k in range(0,G.shape[0],3)])
        self.keep_rows=np.concatenate([np.arange(3*j,3*j+3) for j in self.base.cone_ids])
    def at(self,fs):
        t=time.perf_counter();old=self.base
        p=Prepared.__new__(Prepared);p.__dict__=old.__dict__.copy();p.fs=fs
        E0,b0,G0,h0,dims0,cost0=old.original;G=G0.copy();h=h0.copy()
        st=np.array([audit_strength(self.m,e,fs) for e in range(len(self.m.tri))])
        sine=np.sin(st[:,1]);G.data[self.head_data_mask]=-sine[self.head_data_owners]
        h[::3]=2*st[self.cone_owners,0]*np.cos(st[self.cone_owners,1])
        p.original=E0,b0,G,h,dims0,cost0
        E,b,_,_,dims,cost=old.reduced
        GG=(G@p.T).tocsc() if p.T is not None else G
        GG=GG[self.keep_rows].tocsc();hh=(h+G@p.off)[self.keep_rows]
        p.reduced=E,b,GG,hh,dims,cost
        p.A=sparse.vstack([E,-GG],format='csc');p.rhs=np.r_[b,hh]
        p.assembly_seconds=time.perf_counter()-t;p.transform_seconds=0.
        return p

def reuse_witness(pr,oldx):
    """Rescale only in the reference-only loading supported by P06; then audit.
    Never use this scaling when there is fixed body/traction work.
    """
    if pr.mode=='lower':
        assert np.max(np.abs(pr.m.body[:,:,0]),initial=0)==0
        assert all(np.max(np.abs(v['loads'][:,0]),initial=0)==0 for v in pr.m.boundary.values())
        x=oldx.copy()
        if x[-1]>LOWER_TARGET:x*=LOWER_TARGET/x[-1]
    else:x=oldx.copy()
    t=time.perf_counter();r=evaluate(pr,x)
    r.update(method='reaudit_cached_primal',audit_seconds=time.perf_counter()-t,fs=pr.fs,mode=pr.mode)
    return r,x

def solve_feasibility(pr,overrides=None):
    """Feasible-field search only; infeasibility/solver error returns UNKNOWN.

    Lower pr must already include a fixed reference load. Upper adds an actual
    objective cap, but optimizes zero; its result cannot supply a root value.
    """
    if pr.mode=='upper':
        E,b,G,h,dims,cost=pr.reduced
        cap=UPPER_TARGET-2e-6-pr.objective_offset
        A=sparse.vstack([pr.A,sparse.csc_matrix(cost.reshape(1,-1))],format='csc')
        rhs=np.r_[pr.rhs,cap];cones=[*pr.cones,clarabel.NonnegativeConeT(1)]
    else:A,rhs,cones=pr.A,pr.rhs,pr.cones
    t=time.perf_counter();s=clarabel.DefaultSolver(sparse.csc_matrix((len(pr.cost),len(pr.cost))),np.zeros(len(pr.cost)),A,rhs,cones,settings(overrides))
    build=time.perf_counter()-t;t=time.perf_counter();sol=s.solve();wall=time.perf_counter()-t
    x=pr.restore(sol.x);t=time.perf_counter()
    if str(sol.status) not in ['PrimalInfeasible','AlmostPrimalInfeasible','DualInfeasible','AlmostDualInfeasible']:
        r=evaluate(pr,x)
    else:r={'physical_pass':False,'witness':False,'value':None,'numerical_gate':False,'reason':'NO_PRIMAL_FIELD'}
    au=time.perf_counter()-t
    ray=None
    if str(sol.status) in ['PrimalInfeasible','AlmostPrimalInfeasible']:
        cp=Prepared.__new__(Prepared);cp.__dict__=pr.__dict__.copy()
        if pr.mode=='upper':
            E0,b0,G0,h0,dims0,c0=pr.original;E,b,G,h,dims,cost=pr.reduced
            cp.original=E0,b0,sparse.vstack([G0,-c0.reshape(1,-1)],format='csc'),np.r_[h0,UPPER_TARGET-2e-6],dims0+[1],np.zeros(len(c0))
            cp.reduced=E,b,sparse.vstack([G,-cost.reshape(1,-1)],format='csc'),np.r_[h,UPPER_TARGET-2e-6-pr.objective_offset],dims+[1],np.zeros(len(cost))
            cp.cone_ids=np.r_[pr.cone_ids,len(dims0)]
        z,rest=cp.restore_dual(np.asarray(sol.z));E0,b0,G0,h0,dims0,c0=cp.original
        pr.last_restored_ray=z
        y,zz=z[:len(b0)],z[len(b0):];sep=-float(b0@y+h0@zz)
        stat=float(np.max(np.abs(E0.T@y-G0.T@zz),initial=0));dc=cone_error(zz,dims0)
        ray={'separation':sep,'stationarity':stat,'dual_cone_error':dc,'dual_restore':rest,
             'normalized_stationarity':stat/sep if sep>0 else None,
             'normalized_cone_error':dc/sep if sep>0 else None,
             'verified':bool(sep>1e-8 and stat/sep<=1e-9 and dc/sep<=1e-10),
             'scope':'DISCRETE_FIXED_FS_NUMERICAL_INFEASIBILITY; not continuum instability'}
    r.update(case=pr.m.provenance['case'],fs=pr.fs,mode=pr.mode,elements=len(pr.m.tri),status=str(sol.status),
             method='fixed_load_feasibility' if pr.mode=='lower' else 'capped_upper_feasibility',
             iterations=sol.iterations,solver_seconds=wall,build_seconds=build,audit_seconds=au,
             transform_seconds=pr.transform_seconds,assembly_seconds=pr.assembly_seconds,
             certificate='ONE_SIDED_PRIMAL_WITNESS' if r['witness'] else 'UNKNOWN',
             returns_root_value=False,zero_objective_solver=True,infeasibility_ray=ray,
             discrete_infeasibility_verified=bool(ray and ray['verified']))
    return r,x

class ConstantFactorADMM:
    """Lower-bound stress splitting with a constant matrix across Fs.

    Cone projection uses the orthonormal stress coordinates (p,d,sqrt(2)*tau).
    Unlike an interior-point matrix, this factor does not change with Fs or
    iteration. It is an experimental first-order method, not an HSD port.
    """
    def __init__(self,m,reduce=True,prox=1e-8):
        self.m=m;t=time.perf_counter();self.base=Prepared(m,.5,'lower',reduction='affine' if reduce else 'none')
        E,b,G,h,dims,cost=self.base.reduced
        ne=len(m.tri);n0=18*ne+1;rr=[];cc=[];vv=[]
        for j in range(6*ne):
            i=3*j
            for r,c,v in [(i,i,-1/np.sqrt(2)),(i,i+1,-1/np.sqrt(2)),(i+1,i,1/np.sqrt(2)),(i+1,i+1,-1/np.sqrt(2)),(i+2,i+2,np.sqrt(2))]:
                rr.append(r);cc.append(c);vv.append(v)
        R0=sparse.coo_matrix((vv,(rr,cc)),shape=(18*ne,n0)).tocsc()
        self.R=(R0@self.base.T).tocsc() if self.base.T is not None else R0
        assert np.max(np.abs(self.base.off),initial=0)==0
        self.A=sparse.vstack([E,self.R],format='csc');self.AT=self.A.T.tocsc();self.b=b;self.neq=len(b);self.cost=cost;self.prox=prox
        matrix=(self.AT@self.A+sparse.eye(len(cost),format='csc')*prox).tocsc()
        t0=time.perf_counter();self.factor=splu(matrix,permc_spec='COLAMD');self.factor_seconds=time.perf_counter()-t0
        self.factor_nnz=self.factor.L.nnz+self.factor.U.nnz;self.setup_seconds=time.perf_counter()-t
        self.factor_count=1
    def projection(self,v,a,b):
        out=v.copy();out[:self.neq]=self.b;w=out[self.neq:].reshape(-1,3)
        pp=w[:,0]+b/a;rad=np.hypot(w[:,1],w[:,2]);inside=rad<=a*pp;polar=(pp<=-a*rad)&(~inside)
        middle=(~inside)&(~polar);newp=(pp+a*rad)/(1+a*a)
        w[polar]=0;w[polar,0]=-b[polar]/a[polar]
        w[middle,0]=newp[middle]-b[middle]/a[middle]
        w[middle,1:]*=(a[middle]*newp[middle]/np.maximum(rad[middle],1e-300))[:,None]
        return out
    def run(self,fs,rho=1.,warm=None,max_iter=20000,seconds=12.,check_every=50,target='witness'):
        m=self.m;pr=Prepared(m,fs,'lower',reduction=self.base.kind)
        assert len(pr.cost)==len(self.cost) and np.max(np.abs((pr.T-self.base.T).data),initial=0)==0
        strengths=np.array([audit_strength(m,e,fs) for e in range(len(m.tri))]);a=np.repeat(np.sin(strengths[:,1]),6);b=np.repeat(np.sqrt(2)*strengths[:,0]*np.cos(strengths[:,1]),6)
        if warm is None:x=np.zeros(len(self.cost));u=np.zeros(self.A.shape[0]);s=self.projection(self.A@x,a,b)
        else:
            x=warm['x'].copy();u=warm['u'].copy()*(warm['rho']/rho);s=self.projection(self.A@x+u,a,b)
        t=time.perf_counter();history=[];r={};accepted=False;loops=0;alpha=1.5
        for k in range(1,max_iter+1):
            xn=self.factor.solve(np.asarray(self.AT@(s-u)-self.cost/rho+self.prox*x));v=self.A@xn;rel=alpha*v+(1-alpha)*s
            sn=self.projection(rel+u,a,b);u+=rel-sn;x=xn;s=sn;loops=k
            if k%check_every==0 or k==max_iter:
                full=pr.restore(x);r=evaluate(pr,full)
                dual=rho*u;yp=dual[self.neq:].reshape(-1,3);dval=-float(self.b@dual[:self.neq])+float((b/a)@yp[:,0])
                stat=float(np.max(np.abs(self.cost+self.AT@dual),initial=0));polarerr=float(np.max(yp[:,0]+a*np.hypot(yp[:,1],yp[:,2]),initial=0))
                gap=abs(float(self.cost@x)-dval)
                r.update(dual_objective=dval,dual_stationarity=stat,dual_polar_error=polarerr,dual_gap=gap,
                         numerical_gate=bool(r['physical_pass'] and stat<=1e-7 and polarerr<=1e-8 and gap<=1e-8*(1+abs(float(self.cost@x)))))
                history.append({'iteration':k,'seconds':time.perf_counter()-t,'value':r['value'],'physical_pass':r['physical_pass'],'witness':r['witness'],'dual_stationarity':stat,'gap':gap,'cone_error':r['cone_error']})
                accepted=r['witness'] if target=='witness' else r['numerical_gate']
                if accepted:break
                if time.perf_counter()-t>=seconds:break
        r.update(case=m.provenance['case'],fs=fs,rho=rho,mode='lower',elements=len(m.tri),method='constant_factor_admm',warm=warm is not None,
                 iterations=loops,solver_seconds=time.perf_counter()-t,setup_seconds=self.setup_seconds,factor_seconds=self.factor_seconds,
                 factor_nnz=self.factor_nnz,factor_count=self.factor_count,accepted=accepted,target=target,history=history,
                 stop='AUDITED_WITNESS' if accepted and target=='witness' else 'NUMERICAL_GATE' if accepted else 'BUDGET_OR_ITERATIONS')
        return r,{'x':x,'u':u,'rho':rho},pr.restore(x)
