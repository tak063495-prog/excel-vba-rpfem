"""Scratch original-dual cone/affine projection; diagnostic files, no published bounds."""
from pathlib import Path
import time
import argparse
from scipy.sparse.linalg import lsmr
import experiment_core as core
import accuracy_driver as driver
from original_farkas import solve_original_farkas,audit_original_farkas,TARGET
np,sparse=core.np,core.sparse


def project_cones(vector,dims):
    result=vector.copy();starts=np.r_[0,np.cumsum(dims[:-1])];dd=np.asarray(dims)
    for dimension in np.unique(dd):
        rows=starts[dd==dimension,None]+np.arange(dimension)[None,:]
        values=result[rows]
        if dimension==1:result[rows]=np.maximum(values,0);continue
        head,tail=values[:,0],values[:,1:];norm=np.linalg.norm(tail,axis=1)
        negative=norm<=-head;inside=norm<=head;middle=~negative & ~inside
        values[negative]=0
        scale=.5*(1+head[middle]/norm[middle])
        values[middle,0]=.5*(head[middle]+norm[middle])
        values[middle,1:]*=scale[:,None]
        result[rows]=values
    return result


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--cases',nargs='+',choices=['D01A','D02A'],default=['D01A','D02A'])
    parser.add_argument('--max-iterations',type=int,default=100)
    args=parser.parse_args()
    if not 0<=args.max_iterations<=100:raise ValueError('Projection iteration cap must be 0..100')
    root=Path(__file__).resolve().parents[1]
    output=root/'results'/('farkas_projection_'+str(time.time_ns()))
    output.mkdir()
    probes=[('D01A','d01_final','D01A_C3',.5792724218488321),
            ('D02A','d02_final','D02A_C7',1.50600141054559)]
    for case,source,label,fs in probes:
        if case not in args.cases:continue
        _,rec=driver.load_model(case)
        path=root/'results'/source/case/(label+'_mesh.npz')
        mesh=driver.load_saved_mesh(path,rec)
        pr=core.Prepared(mesh,fs,'upper',8,reduction='none')
        E,b,G,h,dims,cost=pr.original
        assert np.all(cost==0)
        original=solve_original_farkas(pr,seconds=30,return_dual=True)
        d=original.pop('dual');ny=len(b)
        stationary=sparse.hstack([E.T,-G.T],format='csc')
        M=sparse.vstack([stationary,sparse.csc_matrix(np.r_[-b,-h].reshape(1,-1))],format='csc')
        rhs=np.r_[np.zeros(len(cost)),TARGET]
        report={'case':case,'fs':fs,'q':8,'mesh_id':mesh.mesh_id,'physical_model_id':mesh.physical_model_id,
            'mesh_path':str(path),'original_helper':original,'affine_shape':M.shape,'iterations':[]}
        def audit(label,dual,solver=None):
            row={'label':label,'audit':audit_original_farkas(pr,dual),'lsmr':solver}
            report['iterations'].append(row)
            core.write(output/(case+'_audit.json'),report)
            print(case,label,row['audit']['dual_stationarity'],row['audit']['dual_cone_error'],
                  row['audit']['dual_objective'],row['audit']['accepted'],flush=True)
            return row['audit']['accepted']
        audit('before',d)
        d[ny:]=project_cones(d[ny:],dims)
        sep=-float(b@d[:ny]+h@d[ny:])
        if np.isfinite(sep) and sep>0:d*=TARGET/sep
        accepted=audit('cone_projection_only',d)
        for iteration in range(1,args.max_iterations+1):
            if accepted:break
            t=time.perf_counter();correction=lsmr(M,rhs-M@d,atol=1e-15,btol=1e-15,conlim=1e16,maxiter=10000)
            d+=correction[0]
            d[ny:]=project_cones(d[ny:],dims)
            sep=-float(b@d[:ny]+h@d[ny:])
            if np.isfinite(sep) and sep>0:d*=TARGET/sep
            info={'flag':correction[1],'iterations':correction[2],'normr':correction[3],
                  'normar':correction[4],'normA':correction[5],'condition':correction[6],
                  'seconds':time.perf_counter()-t}
            accepted=audit('affine_then_cone_'+str(iteration),d,info)
            if iteration==20:print(case,'20_ITERATION_CHECKPOINT',accepted,flush=True)
        np.savez_compressed(output/(case+'_dual.npz'),dual=d,fs=fs,q=8,mode='upper',
            accepted=accepted,mesh_id=mesh.mesh_id,physical_model_id=mesh.physical_model_id)
        report['accepted']=accepted;core.write(output/(case+'_audit.json'),report)
    print('DIAGNOSTIC',output,flush=True)


if __name__=='__main__':main()
