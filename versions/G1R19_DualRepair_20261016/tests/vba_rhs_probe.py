"""Apply the VBA compensated CSR residual to frozen N; one DLL linear call only."""
from pathlib import Path
import argparse,hashlib,json,os,sys
for k in ['OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS']:os.environ[k]='1'
import native_lsmr_probe as ffi
np,sparse=ffi.np,ffi.sparse
ROOT=Path(__file__).resolve().parents[1];REPO=ROOT.parents[1]


def rhs_like_vba(A,cost,dual):
    rhs=-cost.copy();compensation=np.zeros(len(cost))
    for i in range(A.shape[0]):
        for j in range(A.indptr[i],A.indptr[i+1]):
            col=int(A.indices[j]);value=-float(A.data[j])*float(dual[i])
            adjusted=value-float(compensation[col]);updated=float(rhs[col])+adjusted
            compensation[col]=(updated-float(rhs[col]))-adjusted;rhs[col]=updated
    return rhs


def run(args):
    if args.output.exists():raise ValueError('Fresh output required')
    pipeline=REPO/'validation/G1R19_python26_terminal_final_20261016'
    sys.path.insert(0,str(pipeline/'python'))
    import accuracy_driver as driver
    from original_dual import audit_original_dual
    source_hashes=driver.source_hashes()
    meshpath=pipeline/'results/P06N_immutable_20261016T124645Z/best_upper_mesh.npz'
    dualpath=pipeline/'results/final_guide_20261016T124646Z/P06N_upper8739_final_guide/terminal_original_dual.npz'
    paths=[Path(__file__),Path(ffi.__file__),ffi.DLL,meshpath,dualpath]
    hashes={p:ffi.sha(p) for p in paths}
    _,record=driver.load_model('P06N');mesh=driver.load_saved_mesh(meshpath,record)
    with np.load(dualpath,allow_pickle=False) as data:dual=data['dual'].copy();fs=float(data['fs'])
    prepared=driver.core.Prepared(mesh,fs,'upper',8,reduction='none')
    E,b,G,h,dims,cost=prepared.original;ny=len(b)
    A=sparse.vstack([E,-G],format='csr');A.sort_indices()
    rhs=rhs_like_vba(A,cost,dual)
    numpy_rhs=-(cost+E.T@dual[:ny]-G.T@dual[ny:])
    correction,linear=ffi.solve(E.T.tocsr(),rhs)
    candidate=dual.copy();candidate[:ny]+=correction
    np.testing.assert_array_equal(candidate[ny:].view(np.uint64),dual[ny:].view(np.uint64))
    after=audit_original_dual(prepared,candidate)
    vba_stat=float(np.max(np.abs(rhs_like_vba(A,cost,candidate)),initial=0.))
    if not after['accepted'] or vba_stat>1e-7 or linear['seconds']>10:raise ValueError('Original full dual gate failed')
    if source_hashes!=driver.source_hashes() or any(ffi.sha(p)!=h for p,h in hashes.items()):raise ValueError('Frozen inputs changed')
    result=dict(status='PASS',optimizer_calls=0,method='one_native_lsmr_with_VBA_compensated_original_CSR_rhs',
                before=audit_original_dual(prepared,dual),after=after,linear=linear,
                compensated_before_stat=float(np.max(np.abs(rhs),initial=0.)),compensated_after_stat=vba_stat,
                compensated_vs_numpy_rhs_maxdifference=float(np.max(np.abs(rhs-numpy_rhs),initial=0.)),
                z_bitwise_unchanged=True,source_vector_unchanged=True,original_files_unchanged=True,
                input_sha256={str(p):h for p,h in hashes.items()})
    args.output.mkdir(parents=True)
    np.savez_compressed(args.output/'native_compensated_repaired.npz',dual=candidate,fs=fs,q=8,mode='upper')
    (args.output/'RESULTS.json').write_text(json.dumps(result,indent=2)+'\n',encoding='utf8')
    print(json.dumps(result),flush=True)
if __name__=='__main__':
    p=argparse.ArgumentParser();p.add_argument('--output',type=Path,required=True);run(p.parse_args())
