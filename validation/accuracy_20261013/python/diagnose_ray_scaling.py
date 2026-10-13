"""Re-audit all recovery iterates at recorded UNKNOWN upper points; no fields published."""
from pathlib import Path
import json
import time
import math
import hashlib
import experiment_core as core
import accuracy_driver as driver

PROBES=[('D01A','d01_recovery','D01A_C2',.579209921848832),
        ('D02A','d_combined','D02A_C3',1.5129451354927583)]


def main():
    base=Path(__file__).resolve().parents[1]/'results'
    output=base/('ray_iterate_diagnostic_'+str(time.time_ns())+'.json')
    reports=[]
    options=[(driver.ACCURATE,False),
        (dict(driver.ACCURATE,iterative_refinement_max_iter=30,
              iterative_refinement_abstol=1e-15,iterative_refinement_reltol=1e-15),False),
        (dict(driver.ACCURATE,tol_gap_abs=1e-13,tol_gap_rel=1e-13,tol_feas=1e-13,
              iterative_refinement_max_iter=30,iterative_refinement_abstol=1e-15,
              iterative_refinement_reltol=1e-15),True),
        ({'equilibrate_enable':True},False)]
    source={name:hashlib.sha256(Path(__file__).with_name(name).read_bytes()).hexdigest()
            for name in ('experiment_core.py','accuracy_driver.py','reduction.py')}
    for case,folder,label,fs in PROBES:
        _,rec=driver.load_model(case)
        mesh_path=base/folder/case/(label+'_mesh.npz')
        mesh=driver.load_saved_mesh(mesh_path,rec)
        prepared=core.Prepared(mesh,fs,'upper',8,reduction='affine')
        E,b,G,h,dims,cost=prepared.original
        assert core.np.all(cost==0)
        def metrics(z):
            y,zc=z[:len(b)],z[len(b):]
            return {'objective':-float(b@y+h@zc),
                'stationarity':float(core.np.max(core.np.abs(E.T@y-G.T@zc),initial=0)),
                'cone_error':core.cone_error(zc,dims)}
        for index,(settings,scaled) in enumerate(options):
            result,_,canonical=core.solve(prepared,dict(settings,time_limit=120),max_iter=350,
                cone_scaling=scaled,primal_polish=True)
            z,_=prepared.restore_dual(canonical)
            before=metrics(z)
            factor=(core.UPPER_TARGET+4e-6)/before['objective'] if before['objective']>0 else None
            after=metrics(factor*z) if factor is not None and math.isfinite(factor) and factor>1 else None
            applicable=bool(result['status'] in ('PrimalInfeasible','AlmostPrimalInfeasible') and
                after and all(math.isfinite(v) for v in before.values()) and
                all(math.isfinite(v) for v in after.values()))
            passed=bool(applicable and after['stationarity']<=1e-7 and after['cone_error']<=1e-8 and
                after['objective']>before['objective'])
            row={'case':case,'fs':fs,'mesh_path':str(mesh_path),'mesh_id':mesh.mesh_id,'q':8,
                'configuration':index,'scaled':scaled,'status':result['status'],
                'existing_2x_recovery':result.get('original_dual_ray_rescaling'),
                'before':before,'small_factor':factor,'after':after,'small_factor_audit_pass':passed}
            reports.append(row)
            core.write(output,{'diagnostic_only':True,'source_sha256':source,'reports':reports})
            print(case,index,result['status'],before,'small_factor_pass',passed,flush=True)
    print('DIAGNOSTIC',output,flush=True)


if __name__=='__main__':main()
