"""Audited SOCP field search with SciPy/HiGHS (no optional conic dependency).

outer_cut: iterated supporting halfspaces; accept ONLY after original cone and
physical audit. Inner polygon infeasibility and LP status never certify the
continuum unstable. Numerical statuses are deliberately separate from witnesses.
Affine elimination preserves equations and is checked against the original model.
"""
import os
for key in ('OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS'):os.environ.setdefault(key,'1')
import time,math
from pathlib import Path
import numpy as np
from scipy import sparse
from scipy.optimize import linprog
from common import ROOT,assemble_lower,assemble,assemble_zero,physical_audit,dump,save_mesh

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

def program(m,fs,mode,q):
    if mode=='lower':
        E,b,G,h,dims,_=assemble_lower(m,fs,fixed_load=1.);cost=np.zeros(E.shape[1]);return E,b,G,h,dims,cost
    p=assemble_zero(m,fs,q) if np.all(m.materials[:,1]==0) else assemble(m,fs,q)
    return p.eq.tocsc(),p.rhs,p.G,np.zeros(p.G.shape[0]),p.dims,p.objective

def build_facets(G,h,dims,facets,inner=False):
    heads=[];ones=[];k=0
    for d in dims:
        if d==3:heads.append(k)
        elif d==1:ones.append(k)
        else:raise ValueError('unsupported cone dimension')
        k+=d
    heads=np.array(heads,int);ones=np.array(ones,int);parts=[];rhs=[]
    if len(ones):parts.append(-G[ones]);rhs.append(h[ones])
    if len(heads):
        parts.append(-G[heads]);rhs.append(h[heads]);factor=math.cos(math.pi/facets) if inner else 1.
        for j in range(facets):
            a=2*math.pi*j/facets;co,si=math.cos(a),math.sin(a)
            parts.append(co*G[heads+1]+si*G[heads+2]-factor*G[heads]);rhs.append(factor*h[heads]-co*h[heads+1]-si*h[heads+2])
    return sparse.vstack(parts,format='csc'),np.concatenate(rhs),heads

def cone_error(v,dims):
    k=0;err=0.
    for d in dims:
        err=max(err,(-v[k] if d==1 else np.linalg.norm(v[k+1:k+d])-v[k]));k+=d
    return float(err)

def checked_field(m,fs,mode,q,x,pr):
    E,b,G,h,dims,cost=pr
    if x is None or not np.isfinite(x).all():return {'accepted':False,'reason':'NO_FINITE_FIELD'}
    eq=float(np.max(np.abs(E@x-b),initial=0));con=cone_error(G@x+h,dims)
    au=physical_audit(m,x,mode,fs,q)
    # For phi=0 the epigraph is a conservative bound, not necessarily tight.
    if mode=='upper' and np.all(m.materials[:,1]==0):
        from zero_phi import audit_zero
        true=float(au['value']);ep=float(cost@x)
        au=audit_zero(m,x,fs,q,true);au['epigraph_objective']=ep;au['epigraph_upper_check']=bool(true<=ep+2e-7)
    meaningful=abs(x[-1]-1)<=1e-9 if mode=='lower' else (au['value']<=1-1e-7 and float(cost@x)<=1-1e-7)
    # Original stress/load units recorded, not silently substituted for normalized errors.
    raw={k:float(v*m.stress_scale) for k,v in au['errors'].items()} if mode=='lower' else {}
    accepted=bool(au['pass'] and eq<=1e-7 and con<=1e-7 and meaningful and au.get('epigraph_upper_check',True))
    return {'accepted':accepted,'audit':au,'eq_error':eq,'cone_error':con,'raw_lower_errors':raw,'raw_lower_pass':max(raw.values(),default=0)<=1e-7,'meaningful_bound':bool(meaningful)}

def solve_field(m,fs,mode,q=2,method='outer_cut',facets=8,reduce=True,seconds=60.,max_rounds=30,label='trial',target=1-2e-6,lp_method='highs-ipm',norm_objective=False):
    start=time.perf_counter();original=program(m,fs,mode,q)
    E,b,G,h,dims,cost=original
    result={'case':m.provenance.get('case'),'label':label,'fs':fs,'mode':mode,'q':q,'method':method,'facets':facets,'affine_reduce':reduce,'norm_objective':norm_objective,'element_count':len(m.tri),'original_variables':E.shape[1],'original_equalities':E.shape[0],'original_cones':len(dims),'seconds_limit':seconds,'accepted':False,'full_root_search':False,'scope':'FELA_ASSOCIATED' if m.provenance.get('psi_policy')=='equal_reduced_phi' else 'FELA_DAVIS_EQUIVALENT','rounds':[]}
    x=None
    try:
        E,b,G,h,dims,T,off,st=affine_reduce(E,b,G,h,dims,reduce)
        result.update(reduction=st,reduced_variables=E.shape[1],reduced_equalities=E.shape[0],reduced_cones=len(dims))
        A,rhs,heads=build_facets(G,h,dims,facets,inner=method=='inner')
        if mode=='upper':A=sparse.vstack([A,sparse.csc_matrix((cost@T).reshape(1,-1))],format='csc');rhs=np.r_[rhs,target-cost@off]
        cuts=[];cut_rhs=[]
        for iteration in range(max_rounds if method=='outer_cut' else 1):
            left=seconds-(time.perf_counter()-start)
            if left<=0:result['stop']='TOTAL_BUDGET';break
            AA=sparse.vstack([A,*cuts],format='csc') if cuts else A
            bb=np.r_[rhs,*cut_rhs] if cuts else rhs
            t=time.perf_counter()
            EE=E;obj=np.zeros(E.shape[1])
            if norm_objective:
                nv=E.shape[1]
                AA=sparse.hstack([AA,sparse.csc_matrix((AA.shape[0],1))],format='csc')
                signs=sparse.vstack([sparse.eye(nv),-sparse.eye(nv)],format='csc')
                box=sparse.hstack([signs,-np.ones((2*nv,1))],format='csc')
                AA=sparse.vstack([AA,box],format='csc');bb=np.r_[bb,np.zeros(2*nv)]
                EE=sparse.hstack([E,sparse.csc_matrix((E.shape[0],1))],format='csc');obj=np.r_[obj,1.]
            sol=linprog(obj,A_ub=AA,b_ub=bb,A_eq=EE,b_eq=b,bounds=(None,None),method=lp_method,options={'time_limit':max(.05,left),'primal_feasibility_tolerance':1e-9,'dual_feasibility_tolerance':1e-9,'ipm_optimality_tolerance':1e-10,'threads':1})
            rec={'iteration':iteration,'lp_status':int(sol.status),'lp_message':sol.message,'lp_seconds':time.perf_counter()-t,'inequalities':AA.shape[0]};result['rounds'].append(rec)
            if sol.x is None:
                result['stop']='LP_'+str(sol.status);break
            y=sol.x[:-1] if norm_objective else sol.x
            x=np.asarray(T@y+off).ravel();ch=checked_field(m,fs,mode,q,x,original);result.update(ch);rec.update(eq_error=ch.get('eq_error'),cone_error=ch.get('cone_error'),physical_max=max(ch.get('audit',{}).get('errors',{'none':float('inf')}).values()))
            if ch['accepted']:result['stop']='AUDITED_WITNESS';break
            if method!='outer_cut':result['stop']='INNER_AUDIT_REJECT';break
            v=np.asarray(G@y+h);rad=np.hypot(v[heads+1],v[heads+2]);viol=rad-v[heads]
            ix=np.flatnonzero(viol>1e-10)
            if not len(ix):result['stop']='NONCONE_AUDIT_REJECT';break
            hd=heads[ix];co=v[hd+1]/rad[ix];si=v[hd+2]/rad[ix]
            new=sparse.diags(co)@G[hd+1]+sparse.diags(si)@G[hd+2]-G[hd]
            br=h[hd]-co*h[hd+1]-si*h[hd+2]
            # Tangent support planes only, no approximation presented as a proof.
            cuts.append(new.tocsc());cut_rhs.append(br);rec['cuts_added']=len(ix)
        else:result['stop']='MAX_CUT_ROUNDS'
    except Exception as exc:
        import traceback
        result.update(stop='ERROR',error=type(exc).__name__+': '+str(exc),traceback=traceback.format_exc())
    result['seconds']=time.perf_counter()-start
    path=ROOT/'results'/label.replace('.','p')
    if x is not None:
        save_mesh(Path(str(path)+'.npz'),m,x=x,fs=float(fs),mode=mode,q=q,model_key=m.physical_model_id,accepted=result['accepted'])
        result['field_file']=Path(str(path)+'.npz').name
    dump(Path(str(path)+'.json'),result)
    print(label,result['stop'],result['accepted'],round(result['seconds'],3),result.get('cone_error'),flush=True)
    return result,x
