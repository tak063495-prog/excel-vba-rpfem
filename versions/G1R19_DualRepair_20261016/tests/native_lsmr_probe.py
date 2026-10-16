"""Same existing DLL vs SciPy LSMR, including optional frozen N guide; optimizer0."""
from pathlib import Path
import argparse,ctypes as C,hashlib,json,os,sys,time
for k in ['OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS']:os.environ[k]='1'
import numpy as np
from scipy import sparse
from scipy.sparse.linalg import lsmr
ROOT=Path(__file__).resolve().parents[1];REPO=ROOT.parents[1]
DLL=REPO/'versions/G1R19/delivery/rpfem_native_x64.dll'
I,U,F=C.c_int32,C.c_uint32,C.c_double
PI,PD=C.POINTER(I),C.POINTER(F)
class Problem(C.Structure):
    _fields_=[('size',U),('abi',U),('n',U),('m',U),('nnz',U),('reserved',U),
              ('rows',PI),('cols',PI),('values',PD),('b',PD),('x0',PD)]
class Report(C.Structure):
    _fields_=[('size',U),('abi',U),('stop',U),('iterations',U),('normr',F),
              ('normar',F),('norma',F),('conda',F),('normx',F)]
def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def solve(A,b):
    A=A.tocsr(copy=True);A.sort_indices()
    rows=np.asarray(A.indptr,np.int32);cols=np.asarray(A.indices,np.int32);values=np.asarray(A.data,float)
    b=np.asarray(b,float);zero=np.zeros(A.shape[1]);out=zero.copy()
    pr=Problem(64,1,A.shape[1],A.shape[0],A.nnz,0,rows.ctypes.data_as(PI),cols.ctypes.data_as(PI),
               values.ctypes.data_as(PD),b.ctypes.data_as(PD),zero.ctypes.data_as(PD));report=Report(56,1)
    dll=C.CDLL(str(DLL));fn=dll.rpfem_accuracy_equalities
    fn.restype=I;fn.argtypes=[C.POINTER(Problem),U,F,F,F,PD,U,C.POINTER(Report)]
    assert (C.sizeof(Problem),C.sizeof(Report))==(64,56)
    started=time.perf_counter();rc=fn(C.byref(pr),200,1e-15,1e-15,1e12,out.ctypes.data_as(PD),len(out),C.byref(report))
    seconds=time.perf_counter()-started
    if rc:raise ValueError('Native equality API rc='+str(rc))
    reference=lsmr(A,b,atol=1e-15,btol=1e-15,conlim=1e12,maxiter=200)
    np.testing.assert_allclose(out,reference[0],atol=2e-13,rtol=2e-12)
    return out,dict(returncode=rc,seconds=seconds,n=int(A.shape[1]),m=int(A.shape[0]),nnz=int(A.nnz),
                    stop=int(report.stop),iterations=int(report.iterations),scipy_stop=int(reference[1]),
                    scipy_iterations=int(reference[2]),maximum_correction_difference=float(np.max(np.abs(out-reference[0]),initial=0.)),
                    optimizer_calls=0)
def run(args):
    if args.output.exists():raise ValueError('Fresh output required')
    args.output.mkdir(parents=True)
    tracked={DLL:sha(DLL),Path(__file__):sha(__file__)}
    A=sparse.csr_matrix([[1.,1.],[2.,2.],[3.,3.]])
    correction,toy=solve(A,np.array([1.,2.,4.]))
    np.testing.assert_allclose(correction,[17/28,17/28],atol=2e-14,rtol=0)
    result=dict(synthetic=toy,optimizer_calls=0,dll_sha256=sha(DLL),status='PASS')
    if args.real_n:
        pipeline=REPO/'validation/G1R19_python26_terminal_final_20261016'
        sys.path.insert(0,str(pipeline/'python'))
        import accuracy_driver as driver
        from original_dual import audit_original_dual
        source_hashes=driver.source_hashes()
        meshpath=pipeline/'results/P06N_immutable_20261016T124645Z/best_upper_mesh.npz'
        dualpath=pipeline/'results/final_guide_20261016T124646Z/P06N_upper8739_final_guide/terminal_original_dual.npz'
        tracked.update({p:sha(p) for p in [meshpath,dualpath]})
        _,rec=driver.load_model('P06N');mesh=driver.load_saved_mesh(meshpath,rec)
        with np.load(dualpath,allow_pickle=False) as data:
            source=data['dual'].copy();fs=float(data['fs']);assert int(data['q'])==8 and str(data['mode'])=='upper'
        prepared=driver.core.Prepared(mesh,fs,'upper',8,reduction='none')
        E,b,G,h,dims,cost=prepared.original;ny=len(b)
        before=audit_original_dual(prepared,source)
        rhs=-(cost+E.T@source[:ny]-G.T@source[ny:])
        dy,record=solve(E.T.tocsr(),rhs)
        repaired=source.copy();repaired[:ny]+=dy
        np.testing.assert_array_equal(repaired[ny:].view(np.uint64),source[ny:].view(np.uint64))
        after=audit_original_dual(prepared,repaired)
        if not after['accepted'] or record['seconds']>10:raise ValueError('Native correction failed original raw gate')
        if source_hashes!=driver.source_hashes():raise ValueError('Pipeline source changed')
        np.savez_compressed(args.output/'native_repaired_dual.npz',source=source,dual=repaired,fs=fs,q=8,mode='upper')
        result['real_n']=dict(before=before,after=after,linear=record,fs=fs,equalities=ny,
                              z_bitwise_unchanged=True,no_physical_bound=True)
    for p,h in tracked.items():
        if sha(p)!=h:raise ValueError('Original input/DLL/probe changed')
    result['original_files_unchanged']=True
    result['input_sha256']={str(p):h for p,h in tracked.items()}
    (args.output/'RESULTS.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf8')
    print(json.dumps(result),flush=True)
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--output',type=Path,required=True);p.add_argument('--real-n',action='store_true')
    run(p.parse_args())
