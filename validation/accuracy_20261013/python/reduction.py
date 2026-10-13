"""Presolve functions copied verbatim from the supplied G1R17 validation runtime.
They use numerical tolerances; original-model checks are mandatory.
"""
import numpy as np
from scipy import sparse

def clean_equalities(E,b,tol=2e-11):
    E=E.tocsr();E.eliminate_zeros();active=np.diff(E.indptr)>0
    if np.any(np.abs(b[~active])>tol):raise ValueError('AFFINE_CONTRADICTION')
    return E[active].tocsc(),b[active]

def affine_reduce(E,b,G,h,dims,enabled=True,max_sweeps=12):
    """Only singleton substitution and exact-form two-variable elimination.
    No rank truncation, no positive cohesion rounding. Column transforms have at
    most one nonzero per original variable, plus a known affine offset.
    Identically-zero SOC heads expose the zero face; tails become equalities.
    """
    n=E.shape[1];T=sparse.eye(n,format='csc');offset=np.zeros(n);stats=[];original_cone_ids=np.arange(len(dims),dtype=int)
    if not enabled:return E,b,G,h,dims,T,offset,stats
    for sweep in range(max_sweeps):
        E,b=clean_equalities(E,b);Ec=E.tocsr();n=E.shape[1]
        parent=np.arange(n);weight=np.ones(n);bias=np.zeros(n);fixed={}
        def find(i):
            if parent[i]!=i:
                p=int(parent[i]);r,w,d=find(p)
                bias[i]+=weight[i]*d;weight[i]*=w;parent[i]=r
            return int(parent[i]),float(weight[i]),float(bias[i])
        for r in range(E.shape[0]):
            a,bp=Ec.indptr[r:r+2];ii=Ec.indices[a:bp];vv=Ec.data[a:bp]
            if len(ii)==1 and abs(vv[0])>=1e-10:
                root,w,d=find(int(ii[0]));value=(float(b[r])/vv[0]-d)/w
                if root in fixed and abs(fixed[root]-value)>2e-10:raise ValueError('CONFLICTING_SINGLETON')
                fixed[root]=value
        merged=0
        # Disjoint homogeneous pairs only: avoids unchecked numerical rank decisions.
        for r in range(E.shape[0]):
            a,bp=Ec.indptr[r:r+2];ii=Ec.indices[a:bp];vv=Ec.data[a:bp]
            if len(ii)!=2 or b[r]!=0 or max(abs(vv))<1e-10:continue
            i,wi,di=find(int(ii[0]));j,wj,dj=find(int(ii[1]))
            if i==j or i in fixed or j in fixed or di!=0 or dj!=0:continue
            ai,aj=vv[0]*wi,vv[1]*wj
            if min(abs(ai),abs(aj))<0.05*max(abs(ai),abs(aj)):continue
            # largest coefficient eliminated -> bounded multiplier
            if abs(ai)>=abs(aj):parent[i]=j;weight[i]=-aj/ai
            else:parent[j]=i;weight[j]=-ai/aj
            merged+=1
        roots=[];ww=[];dd=[]
        for i in range(n):
            r,w,d=find(i)
            roots.append(r);ww.append(w);dd.append(d)
        active=sorted(set(roots)-set(fixed));lookup={r:i for i,r in enumerate(active)}
        rr=[];cc=[];data=[];off=np.zeros(n)
        for i,(r,w,d) in enumerate(zip(roots,ww,dd)):
            off[i]=d
            if r in fixed:off[i]+=w*fixed[r]
            else:rr.append(i);cc.append(lookup[r]);data.append(w)
        S=sparse.coo_matrix((data,(rr,cc)),shape=(n,len(active))).tocsc()
        changed=len(active)<n
        if changed:
            b=b-E@off;h=h+G@off;E=(E@S).tocsc();G=(G@S).tocsc()
            offset=offset+T@off;T=(T@S).tocsc()
        G.eliminate_zeros();gc=G.tocsr();keep=[];newdims=[];kept_cone_ids=[];extra=[];erhs=[];k=0;faces=0
        for cone_idx,dim in enumerate(dims):
            if dim>1 and gc.getrow(k).nnz==0 and h[k]==0:
                for j in range(1,dim):
                    row=gc.getrow(k+j)
                    if row.nnz or h[k+j]!=0:extra.append(row);erhs.append(-h[k+j])
                faces+=1
            elif gc[k:k+dim].nnz==0:
                if (dim==1 and h[k]<-1e-13) or(dim>1 and np.linalg.norm(h[k+1:k+dim])>h[k]+1e-13):raise ValueError('CONSTANT_CONE_CONTRADICTION')
            else:keep.extend(range(k,k+dim));newdims.append(dim);kept_cone_ids.append(original_cone_ids[cone_idx])
            k+=dim
        if extra:E=sparse.vstack([E,*extra],format='csc');b=np.r_[b,erhs]
        G=G[keep];h=h[keep];dims=newdims;original_cone_ids=np.asarray(kept_cone_ids,dtype=int)
        stats.append({'sweep':sweep,'variables_before':n,'variables_after':len(active),'singletons':len(fixed),'pairs':merged,'zero_faces':faces})
        if not changed and not extra:break
    E,b=clean_equalities(E,b)
    if stats:stats[-1]['retained_original_cones']=original_cone_ids.tolist()
    return E,b,G,h,dims,T,offset,stats

def cone_error(v,dims):
    if not np.isfinite(v).all():return float('inf')
    k=0;err=0.
    for d in dims:
        err=max(err,(-v[k] if d==1 else np.linalg.norm(v[k+1:k+d])-v[k]));k+=d
    return float(err)


def affine_reduce_tracked(E,b,G,h,dims,cost,max_sweeps=12):
    """Only singleton substitution and exact-form two-variable elimination.
    No rank truncation, no positive cohesion rounding. Column transforms have at
    most one nonzero per original variable, plus a known affine offset.
    Identically-zero SOC heads expose the zero face; tails become equalities.
    """
    n=E.shape[1];T=sparse.eye(n,format='csc');offset=np.zeros(n);stats=[];original_cone_ids=np.arange(len(dims),dtype=int)
    history=[];current_cost=np.asarray(cost).copy()
    for sweep in range(max_sweeps):
        E.eliminate_zeros();input_keep=np.flatnonzero(np.asarray(E.getnnz(axis=1)).ravel()>0)
        E,b=clean_equalities(E,b);Ec=E.tocsr();n=E.shape[1]
        preE,preG,pre_cost,preT=E.copy(),G.copy(),current_cost.copy(),T.copy();pivots=[]
        parent=np.arange(n);weight=np.ones(n);bias=np.zeros(n);fixed={}
        def find(i):
            if parent[i]!=i:
                p=int(parent[i]);r,w,d=find(p)
                bias[i]+=weight[i]*d;weight[i]*=w;parent[i]=r
            return int(parent[i]),float(weight[i]),float(bias[i])
        for r in range(E.shape[0]):
            a,bp=Ec.indptr[r:r+2];ii=Ec.indices[a:bp];vv=Ec.data[a:bp]
            if len(ii)==1 and abs(vv[0])>=1e-10:
                root,w,d=find(int(ii[0]));value=(float(b[r])/vv[0]-d)/w
                if root in fixed and abs(fixed[root]-value)>2e-10:raise ValueError('CONFLICTING_SINGLETON')
                if root not in fixed:pivots.append(r)
                fixed[root]=value
        merged=0
        # Disjoint homogeneous pairs only: avoids unchecked numerical rank decisions.
        for r in range(E.shape[0]):
            a,bp=Ec.indptr[r:r+2];ii=Ec.indices[a:bp];vv=Ec.data[a:bp]
            if len(ii)!=2 or b[r]!=0 or max(abs(vv))<1e-10:continue
            i,wi,di=find(int(ii[0]));j,wj,dj=find(int(ii[1]))
            if i==j or i in fixed or j in fixed or di!=0 or dj!=0:continue
            ai,aj=vv[0]*wi,vv[1]*wj
            if min(abs(ai),abs(aj))<0.05*max(abs(ai),abs(aj)):continue
            # largest coefficient eliminated -> bounded multiplier
            if abs(ai)>=abs(aj):parent[i]=j;weight[i]=-aj/ai
            else:parent[j]=i;weight[j]=-ai/aj
            merged+=1;pivots.append(r)
        roots=[];ww=[];dd=[]
        for i in range(n):
            r,w,d=find(i)
            roots.append(r);ww.append(w);dd.append(d)
        active=sorted(set(roots)-set(fixed));lookup={r:i for i,r in enumerate(active)}
        rr=[];cc=[];data=[];off=np.zeros(n)
        for i,(r,w,d) in enumerate(zip(roots,ww,dd)):
            off[i]=d
            if r in fixed:off[i]+=w*fixed[r]
            else:rr.append(i);cc.append(lookup[r]);data.append(w)
        S=sparse.coo_matrix((data,(rr,cc)),shape=(n,len(active))).tocsc()
        changed=len(active)<n
        if changed:
            b=b-E@off;h=h+G@off;E=(E@S).tocsc();G=(G@S).tocsc()
            offset=offset+T@off;T=(T@S).tocsc();current_cost=np.asarray(current_cost@S).ravel()
        G.eliminate_zeros();gc=G.tocsr();keep=[];newdims=[];kept_cone_ids=[];extra=[];erhs=[];k=0;faces=0
        for cone_idx,dim in enumerate(dims):
            if dim>1 and gc.getrow(k).nnz==0 and h[k]==0:
                for j in range(1,dim):
                    row=gc.getrow(k+j)
                    if row.nnz or h[k+j]!=0:extra.append(row);erhs.append(-h[k+j])
                faces+=1
            elif gc[k:k+dim].nnz==0:
                if (dim==1 and h[k]<-1e-13) or(dim>1 and np.linalg.norm(h[k+1:k+dim])>h[k]+1e-13):raise ValueError('CONSTANT_CONE_CONTRADICTION')
            else:keep.extend(range(k,k+dim));newdims.append(dim);kept_cone_ids.append(original_cone_ids[cone_idx])
            k+=dim
        if extra:raise ValueError("TRACKED_FACE_DUAL_NOT_IMPLEMENTED")
        G=G[keep];h=h[keep];dims=newdims;original_cone_ids=np.asarray(kept_cone_ids,dtype=int)
        history.append({'E':preE,'G':preG,'cost':pre_cost,'T':preT,'input_keep':input_keep,'cone_keep':np.asarray(keep,int),'pivots':np.asarray(pivots,int),'roots':np.asarray(active,int)})
        stats.append({'sweep':sweep,'variables_before':n,'variables_after':len(active),'singletons':len(fixed),'pairs':merged,'zero_faces':faces})
        if not changed and not extra:break
    E.eliminate_zeros();final_keep=np.flatnonzero(np.asarray(E.getnnz(axis=1)).ravel()>0)
    E,b=clean_equalities(E,b)
    if stats:stats[-1]['retained_original_cones']=original_cone_ids.tolist()
    return E,b,G,h,dims,T,offset,stats,history,final_keep
