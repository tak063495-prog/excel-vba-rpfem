"""Analytic phi=0 checks, benchmark clay solves, and frozen phi>0 regression.

Run with the experiment Python environment; no Excel or DLL is required.
"""
from pathlib import Path
import importlib.util
import json
import sys

import experiment_core as core
from experiment_core import np
from model_support.schema import Mesh,topology
from model_support.p2 import gradients
from model_support.upper_assembly import assemble,zero_phi_quadrature,complete_upper_velocity
from independent_audit import audit
from zero_phi_transport import transport,PhiZeroUpperRefresh

ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[2]
FROZEN=ROOT/'data/frozen_phi_positive'

def specimen():
    xy=np.array([[0,0],[1,0],[1,1],[0,1]],float)
    tri=np.array([[0,1,2],[0,2,3]],np.int32)
    boundary={}
    for edge,kind,traction in [((0,1),'fixed',[0,0]),((1,2),'load',[0,1]),
                              ((2,3),'load',[1,0]),((0,3),'load',[0,-1])]:
        boundary[edge]={'kind':kind,'rigid':-1,'loads':np.array([[0,traction[0]],[0,traction[1]]],float)}
    return Mesh(xy,tri,np.array([[2,0,0,0]],float),np.zeros(2,np.int32),
                np.full(2,-1),np.zeros((2,2,2)),boundary,2.,'analytic_uniform_shear',
                provenance={'case':'analytic_phi_zero','psi_policy':'equal_reduced_phi'}).validate()

def solve_case(m,fs,q):
    pr=core.Prepared(m,fs,'upper',q,reduction='none')
    result,x,_=core.solve(pr,{'equilibrate_enable':True})
    assert result['physical_pass'],result
    return {'case':m.provenance['case'],'elements':len(m.tri),'fs':fs,'q':q,
            'value':result['value'],'objective':result['objective'],
            'physical_pass':result['physical_pass'],'numerical_gate':result['numerical_gate'],
            'errors':result['audit']['errors'],'seconds':result['solver_seconds']}

def analytic_checks():
    m=specimen();rows=[]
    # Constant pure shear has exact dissipation c*area*|du/dy|. Both the
    # volume and interface estimates are exact here for every subdivision.
    for q in [1,2,8]:
        r=solve_case(m,2.,q)
        assert abs(r['value']-1.)<2e-7,r
        rows.append(r)
    lr,_,_=core.solve(core.Prepared(m,2.,'lower',8,reduction='none'))
    assert lr['physical_pass'] and abs(lr['value']-1.)<2e-7,lr
    fixed=specimen()
    for value in fixed.boundary.values():value['loads'][:,0]=.2*value['loads'][:,1]
    r=solve_case(fixed,2.,2);assert abs(r['value']-.8)<2e-7,r
    # Independently integrate a nonconstant affine strain norm using a Duffy
    # Gauss rule. The conservative estimate must decrease under nested splits.
    area,grad,_=topology(m);bounds=[]
    g,w=np.polynomial.legendre.leggauss(40);g=(g+1)/2;w=w/2
    exact=0.
    for e,t in enumerate(m.tri):
        p=m.xy[t];points=np.r_[p,(p+np.roll(p,-1,axis=0))/2]
        xx,yy=points.T;v=np.column_stack([xx*xx+2*xx*yy,-2*xx*yy-yy*yy])
        for u,wu in zip(g,w):
            for z,wz in zip(g,w):
                dv=v.T@gradients([1-u,u*(1-z),u*z],grad[e])
                exact+=area[e]*2*u*wu*wz*np.hypot(dv[0,0]-dv[1,1],dv[0,1]+dv[1,0])
    for q in [1,2,4,8]:
        val=0.
        for e,t in enumerate(m.tri):
            p=m.xy[t];points=np.r_[p,(p+np.roll(p,-1,axis=0))/2]
            xx,yy=points.T;v=np.column_stack([xx*xx+2*xx*yy,-2*xx*yy-yy*yy])
            for bary,weight in zero_phi_quadrature(q):
                dv=v.T@gradients(bary,grad[e]);assert abs(np.trace(dv))<1e-12
                val+=area[e]*weight*np.hypot(dv[0,0]-dv[1,1],dv[0,1]+dv[1,0])
        assert val>=exact-1e-9,(q,val,exact)
        if bounds:assert val<=bounds[-1]['conservative_norm_integral']+1e-12
        bounds.append({'q':q,'conservative_norm_integral':float(val),'gauss_reference':float(exact)})
    # A deliberately invalid epigraph is caught by the physical auditor.
    pr=core.Prepared(m,2.,'upper',2,reduction='none');_,x,_=core.solve(pr)
    filled=complete_upper_velocity(m,x[:12*len(m.tri)],2)
    check=core.evaluate(pr,filled);assert check['physical_pass'],check
    bad=x.copy();bad[12*len(m.tri)]-=.1
    result=audit(m,bad,'upper',2.,2,float(pr.original[-1]@bad),18*len(m.tri))
    assert not result['pass'] and result['errors']['epigraph']>.09,result
    return {'uniform_shear_solves':rows,'fixed_traction_shear':r,
            'uniform_shear_lower_value':lr['value'],
            'nonconstant_norm_bounds':bounds,'invalid_epigraph_rejected':True,
            'velocity_epigraph_reconstruction_pass':True}

def clay_mesh(case):
    path=REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs'/f'{case}.json'
    rec=json.loads(path.read_text(encoding='utf-8'))
    nodes=[];nodeid={}
    # This mesh follows the published 45 degree clay slope and y=10 interface.
    # No failure line or material property is changed.
    for y in range(0,21,5):
        xmax=60 if y<=10 else 60-2*y
        for x in range(0,xmax+1,10):nodeid[x,y]=len(nodes);nodes.append([x,y])
    tri=[];ids=[]
    for y in range(0,20,5):
        xmax=60 if y<10 else 60-2*y
        uppermax=60 if y+5<=10 else 60-2*(y+5)
        for x in range(0,xmax,10):
            a=nodeid[x,y];b=nodeid[x+10,y];c=nodeid[x,y+5]
            if x+10<=uppermax:
                d=nodeid[x+10,y+5];tri.extend([[a,b,d],[a,d,c]]);ids.extend([1 if y<10 else 0]*2)
            else:tri.append([a,b,c]);ids.append(0)
    return core.mesh_from_arrays(np.array(nodes,float),np.array(tri,np.int32),np.array(ids,np.int32),rec)

def friction_regression():
    spec=importlib.util.spec_from_file_location('frozen_upper',FROZEN/'upper_assembly.py')
    # Relative imports in the frozen module refer to unchanged support files.
    module=importlib.util.module_from_spec(spec);module.__package__='model_support';sys.modules[spec.name]=module;spec.loader.exec_module(module)
    spec2=importlib.util.spec_from_file_location('frozen_audit',FROZEN/'independent_audit.py')
    oldaudit=importlib.util.module_from_spec(spec2);spec2.loader.exec_module(oldaudit)
    rows=[]
    for case in ['P06A','P06N']:
        m,rec=core.load_case(case)
        d=np.load(FROZEN/f'{case}_upper_reference.npz');x=d['x'];fs=float(d['fs'])
        old=module.assemble(m,fs,8);new=assemble(m,fs,8)
        for name in ['eq','G']:
            a=getattr(old,name).tocsc();b=getattr(new,name).tocsc()
            assert a.shape==b.shape and np.array_equal(a.indptr,b.indptr) and np.array_equal(a.indices,b.indices) and np.array_equal(a.data,b.data),name
        for name in ['rhs','objective','work','row_norm']:assert np.array_equal(getattr(old,name),getattr(new,name)),name
        assert old.dims==new.dims
        a=oldaudit.audit(m,x,'upper',fs,8,float(old.objective@x),18*len(m.tri))
        b=audit(m,x,'upper',fs,8,float(new.objective@x),18*len(m.tri))
        assert a==b,(a,b)
        child=core.uniform_refine(m,rec)
        assert np.array_equal(child.xy,d['child_xy']) and np.array_equal(child.tri,d['child_tri']) and np.array_equal(child.material_id,d['child_material_id'])
        assert np.array_equal(d['prolonged_x'],core.prolong(m,child,x,'upper'))
        rows.append({'case':case,'matrices_bitwise_identical':True,'audits_identical':True,
                     'prolongation_bitwise_identical':True})
    return rows

def prolonged_clay_check():
    m=clay_mesh('P09A')
    rec=json.loads((REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs/P09A.json').read_text(encoding='utf-8'))
    p=core.Prepared(m,1.,'upper',8,reduction='none');r,x,_=core.solve(p)
    score=core.indicator(m,x)
    assert np.array_equal(score,core.indicator(m,x[:12*len(m.tri)]))
    child=core.uniform_refine(m,rec);xp=core.prolong(m,child,x,'upper')
    check=core.evaluate(core.Prepared(child,1.,'upper',8,reduction='none'),xp)
    assert check['physical_pass'] and check['value']<=r['value']+1e-7,check
    return {'case':'P09A','parent_elements':len(m.tri),'child_elements':len(child.tri),
            'velocity_variables':12*len(child.tri),'auxiliary_variables':len(xp)-12*len(child.tri),
            'physical_pass':check['physical_pass'],'indicator_ignores_auxiliary':True,
            'parent_conservative_value':r['value'],'child_conservative_value':check['value']}

def scaling_transport_checks():
    m=specimen();rows=[]
    for side in ['lower','upper']:
        r,x,_=core.solve(core.Prepared(m,1.,side,2,reduction='none'))
        assert r['physical_pass']
        moved,field=transport(m,1.,x,side,2)
        expected=2/(1.0000012 if side=='lower' else .9999988)
        assert abs(moved['fs']-expected)<2e-7 and not moved['root_complete'],moved
        rows.append(moved)
    fixed=specimen()
    for value in fixed.boundary.values():value['loads'][:,0]=.2*value['loads'][:,1]
    rejected=[]
    for label,model in [('fixed_load',fixed),('nonzero_phi',core.load_case('P06A')[0])]:
        try:transport(model,1.,np.array([]),'upper',2)
        except ValueError:rejected.append(label)
        else:raise AssertionError(label+' must be rejected before scaling')
    benchmarks=[]
    for case in ['P08A','P09A','P10A']:
        m=clay_mesh(case)
        for side in ['lower','upper']:
            r,x,_=core.solve(core.Prepared(m,1.,side,2,reduction='none'))
            assert r['physical_pass'],r
            moved,_=transport(m,1.,x,side,2)
            benchmarks.append({'case':case,**moved})
    return {'analytic_shear_certificates':rows,'benchmark_certificates':benchmarks,
            'ineligible_models_rejected':rejected,'discrete_root_optimality_claimed':False}

def upper_refresh_checks():
    m=clay_mesh('P09A');cache=PhiZeroUpperRefresh(m,1.,2,reduction='affine');rows=[]
    for fs in [.6,1.,1.4]:
        refreshed=cache.at(fs);fresh=core.Prepared(m,fs,'upper',2,reduction='affine')
        for group in ['original','reduced']:
            a=getattr(refreshed,group);b=getattr(fresh,group)
            for k in [0,2]:assert (a[k]-b[k]).nnz==0
            for k in [1,3,5]:assert np.allclose(a[k],b[k],rtol=1e-14,atol=1e-14)
            assert a[4]==b[4]
        r,_,_=core.solve(refreshed);s,_,_=core.solve(fresh)
        assert r['physical_pass'] and s['physical_pass'] and abs(r['value']-s['value'])<2e-7,(r,s)
        rows.append({'fs':fs,'value':r['value'],'fresh_value':s['value'],
                     'physical_pass':True,'prepared_refresh_seconds':refreshed.assembly_seconds,
                     'fresh_assembly_transform_seconds':fresh.assembly_seconds+fresh.transform_seconds})
    return rows

def main():
    results={'analytic':analytic_checks(),'friction_regression':friction_regression(),
             'clay_benchmark_fixed_fs':[solve_case(clay_mesh(c),1.,2) for c in ['P08A','P09A','P10A']],
             'prolongation':prolonged_clay_check(),'strength_scaling_transport':scaling_transport_checks(),
             'upper_prepared_refresh':upper_refresh_checks()}
    path=ROOT.parent/'results'/'zero_phi_checks.json';core.write(path,results)
    print(json.dumps(results,ensure_ascii=False,indent=2))

if __name__=='__main__':main()
