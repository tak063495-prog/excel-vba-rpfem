"""Same-DLL, same literal AC18 LP setting comparison; no bound publication.

Only solver settings vary. The saved native physical model is reconstructed
from exact hex doubles, verified against canonical D08 hydrostatic loads, and
audited independently at q8. Original vectors and LP bytes are never modified.
"""
from pathlib import Path
import argparse, ast, ctypes as C, hashlib, importlib.util, json, os, struct, sys, time

ROOT=Path(__file__).resolve().parent
REPO=ROOT.parents[1]
for name in ['OPENBLAS_NUM_THREADS','OMP_NUM_THREADS','MKL_NUM_THREADS']:
    os.environ[name]='1'
sys.path.insert(0,str(REPO/'validation/G1R19_python26_water_20261016/python'))
import experiment_core as core
np,sparse=core.np,core.sparse


def sha(path):return hashlib.sha256(Path(path).read_bytes()).hexdigest()
def require(condition,message):
    if not condition:raise ValueError(message)
def exact(value):return struct.unpack('<d',bytes.fromhex(value))[0]
def decode(raw):return raw.decode('utf16' if raw[:2] in [b'\xff\xfe',b'\xfe\xff'] else 'utf8')


def parse(path):
    raw=Path(path).read_bytes();require(raw[:8]==b'AC18P001','Wrong LP magic');offset=8
    def take(dtype,count):
        nonlocal offset
        a=np.frombuffer(raw,dtype=dtype,count=count,offset=offset).copy();offset+=a.nbytes;return a
    n,m,nz,nc,neq=map(int,take('<i4',5))
    ptr,ids,val=take('<i4',m+1),take('<i4',nz),take('<f8',nz)
    rhs,cost=take('<f8',m),take('<f8',n)
    cones=take('<i4',2*nc).reshape(nc,2)
    fs=float(take('<f8',1)[0]);q=int(take('<i4',1)[0])
    dual='.accuracy_dual_' in path.name
    saved=take('<f8',m if dual else n)
    require(offset==len(raw) and q==8 and not dual,'Primal original q8 export required')
    require(ptr[0]==0 and ptr[-1]==nz and np.all(np.diff(ptr)>=0),'Invalid CSR')
    require(np.array_equal(cones[0],[0,neq]) and int(cones[:,1].sum())==m,'Invalid cones')
    require(all(np.isfinite(a).all() for a in [val,rhs,cost,saved]),'Nonfinite input')
    context={}
    for line in decode(path.with_suffix(path.suffix+'.context.tsv').read_bytes()).splitlines():
        fields=line.split('\t')
        if len(fields)==2 and fields[0] in ['model','fs','q']:context[fields[0]]=fields[1]
    require(exact(context['fs'])==fs and int(context['q'])==q,'Fs/q context differs')
    return dict(n=n,m=m,nz=nz,neq=neq,ptr=ptr,ids=ids,val=val,b=rhs,cost=cost,cones=cones,
                fs=fs,q=q,saved=saved,context=context,A=sparse.csr_matrix((val,ids,ptr),shape=(m,n)))


def model(lp):
    rec=json.loads((REPO/'docs/benchmarks/G1R17_native26_20261013/data/inputs/D08A.json').read_text(encoding='utf8'))
    tokens=lp['context']['model'].split('|');h=tokens[0].split(',')
    nn,ne,nm,nr,ne_edge=map(int,h[:5]);scale=exact(h[7])
    require(nr==0 and h[5:7]==['fixed','reference_total'] and h[8:]==['ASSOCIATED','raw_audit=True'],'Unsupported model context')
    cursor=1;xy=np.array([[exact(v[:16]),exact(v[16:])] for v in tokens[cursor:cursor+nn]]);cursor+=nn
    cells=[];body=[]
    for row in tokens[cursor:cursor+ne]:
        fields=row.split(',');cells.append(list(map(int,fields[:5])))
        body.append([[exact(v[:16]),exact(v[16:])] for v in fields[5:]])
    cursor+=ne;cells=np.array(cells,np.int32);body=np.array(body)
    mats=np.array([[exact(v[i:i+16]) for i in range(0,64,16)] for v in tokens[cursor:cursor+nm]]);cursor+=nm
    boundary={};edge_keys=set()
    for row in tokens[cursor:cursor+ne_edge]:
        fields=row.split(',');a,b,e,other=map(int,fields[:4]);key=tuple(sorted((a,b)));edge_keys.add(key)
        if other<0:
            boundary[key]=dict(kind=fields[4],rigid=int(fields[5]),loads=np.array([[exact(v[:16]),exact(v[16:])] for v in fields[6:]]))
    cursor+=ne_edge
    for row in tokens[cursor:]:
        if row:
            key,a,b=row.split(',');require(tuple(sorted((int(a),int(b)))) in edge_keys,'Constraint edge missing')
    require(np.all(cells[:,4]==-1) and lp['n']==12*ne,'Upper velocity-only original layout required')
    require(np.array_equal(mats,np.array([[r[k] for k in ['c','phi','psi','gamma']] for r in rec['source_case']['materials']])),'Canonical materials differ')
    mesh=core.Mesh(xy,cells[:,:3],mats,cells[:,3],cells[:,4],body,boundary,scale,
                   hashlib.sha256(lp['context']['model'].encode()).hexdigest(),
                   provenance={'psi_policy':rec['source_case']['policy'],'case':'D08A'}).validate()
    y=xy[mesh.tri,1];require(not np.any((y.min(axis=1)<3)&(y.max(axis=1)>3)),'Mesh crosses head3')
    expected=np.zeros_like(body);expected[:,1,1]=-mats[cells[:,3],3]+np.where(y.max(axis=1)<=3,9.81,0.)
    np.testing.assert_allclose(body,expected,atol=1e-12,rtol=0)
    spec=importlib.util.spec_from_file_location('canonical_boundary_ref',REPO/'versions/G1R19_WaterCommon_20261016/tests/test_water_common_geometry.py')
    ref=importlib.util.module_from_spec(spec);spec.loader.exec_module(ref)
    want=ref.canonical_boundary(xy,mesh.tri,rec)
    require(want.keys()==boundary.keys(),'Canonical exterior boundary differs')
    for key,v in want.items():
        require((boundary[key]['kind'],boundary[key]['rigid'])==v[:2],'Boundary restraint differs')
        np.testing.assert_array_equal(boundary[key]['loads'],np.array(v[2:]).reshape(2,2).T)
    area,_,_=core.topology(mesh);force=area@body[:,:,1]
    np.testing.assert_allclose(force,[0.,-2571.4],atol=1e-8,rtol=0)
    return mesh,dict(nodes=nn,elements=ne,area=float(area.sum()),force=force.tolist(),
                     canonical_material_boundary_and_common_hydrostatic_load=True)


def abi(dll_path):
    # Reuse the already validated public C ABI declarations only. Importing the
    # old numerical test module would also import its unrelated model driver.
    path=REPO/'versions/G1R18/tests/test_accuracy_native_ffi.py'
    tree=ast.parse(path.read_text(encoding='utf8'))
    classes=[n for n in tree.body if isinstance(n,ast.ClassDef) and n.name in ['Problem','Settings','Progress']]
    I32,U32,U64,F64=C.c_int32,C.c_uint32,C.c_uint64,C.c_double
    ns=dict(C=C,I32=I32,U32=U32,U64=U64,F64=F64,PI=C.POINTER(I32),PD=C.POINTER(F64))
    exec(compile(ast.Module(body=classes,type_ignores=[]),str(path),'exec'),ns)
    Problem,Settings,Progress=[ns[name] for name in ['Problem','Settings','Progress']]
    require([C.sizeof(t) for t in [Problem,Settings,Progress]]==[88,128,80],'ABI layout differs')
    dll=C.CDLL(str(dll_path))
    sig={'abi_version':(U32,[]),'default_settings':(I32,[C.POINTER(Settings)]),
         'create':(I32,[C.POINTER(Problem),C.POINTER(Settings),C.POINTER(U64)]),
         'start':(I32,[U64]),'poll':(I32,[U64,C.POINTER(Progress)]),'cancel':(I32,[U64]),
         'result':(I32,[U64,ns['PD'],U32,ns['PD'],U32,C.POINTER(Progress)]),'destroy':(I32,[U64])}
    for name,(result,args) in sig.items():
        fn=getattr(dll,'rpfem_accuracy_'+name);fn.restype=result;fn.argtypes=args
    require(dll.rpfem_accuracy_abi_version()==1,'Wrong DLL ABI')
    return dll,ns


def metrics(lp,x,z):
    slack=lp['b']-lp['A']@x;offset=lp['neq'];pc=dc=0.
    for kind,dim in lp['cones'][1:]:
        s=slack[offset:offset+dim];d=z[offset:offset+dim]
        require((kind,dim) in [(1,1),(3,3)],'Unsupported cone')
        pc=max(pc,float(-s[0]+(np.linalg.norm(s[1:]) if kind==3 else 0.)))
        dc=max(dc,float(-d[0]+(np.linalg.norm(d[1:]) if kind==3 else 0.)));offset+=dim
    eq=float(np.max(np.abs(slack[:lp['neq']]),initial=0.));stat=float(np.max(np.abs(lp['cost']+lp['A'].T@z),initial=0.))
    return dict(original_equality=eq,original_primal_cone=pc,original_dual_stationarity=stat,
                original_dual_cone=dc,original_objective=float(lp['cost']@x),
                original_dual_objective=float(-lp['b']@z),
                literal_primal_gate=eq<=1e-7 and pc<=1e-7,
                literal_dual_gate=stat<=1e-7 and dc<=1e-8)


def solve(dll,ns,lp,equ,tol,seconds):
    arrays=dict(row_ptr=lp['ptr'],col_index=lp['ids'],A_values=lp['val'],b=lp['b'],
                Pdiag=np.zeros(lp['n']),q=lp['cost'],cone_kind=lp['cones'][:,0].copy(),cone_dim=lp['cones'][:,1].copy())
    problem=ns['Problem'](struct_size=88,abi_version=1,n=lp['n'],m=lp['m'],nnz=lp['nz'],cone_count=len(lp['cones']))
    for name,a in arrays.items():setattr(problem,name,a.ctypes.data_as(ns['PI'] if a.dtype==np.int32 else ns['PD']))
    settings=ns['Settings']();require(dll.rpfem_accuracy_default_settings(C.byref(settings))==0,'Defaults failed')
    settings.time_limit=seconds;settings.max_threads=1;settings.equilibrate=equ
    for name in ['tol_gap_abs','tol_gap_rel','tol_feas','tol_infeas_abs','tol_infeas_rel']:setattr(settings,name,tol)
    handle=C.c_uint64();require(dll.rpfem_accuracy_create(C.byref(problem),C.byref(settings),C.byref(handle))==0,'Create failed')
    begin=time.monotonic();cancelled=False
    try:
        require(dll.rpfem_accuracy_start(handle)==0,'Start failed')
        while True:
            progress=ns['Progress'](struct_size=80,abi_version=1)
            require(dll.rpfem_accuracy_poll(handle,C.byref(progress))==0,'Poll failed')
            if progress.state not in [0,1]:break
            if time.monotonic()-begin>seconds+5:
                require(dll.rpfem_accuracy_cancel(handle)==0,'Cancel failed');cancelled=True
            if time.monotonic()-begin>seconds+15:raise TimeoutError('Worker failed to cancel')
            time.sleep(.02)
        x,z=np.zeros(lp['n']),np.zeros(lp['m'])
        require(progress.state==2 and not cancelled,'Worker did not complete within budget')
        require(dll.rpfem_accuracy_result(handle,x.ctypes.data_as(ns['PD']),len(x),z.ctypes.data_as(ns['PD']),len(z),C.byref(progress))==0,'Result failed')
        return x,z,dict(equilibrate=equ,tolerance=tol,time_limit=seconds,threads=1,
                        wall_seconds=time.monotonic()-begin,
                        progress={name:getattr(progress,name) for name,_ in progress._fields_},
                        full_settings={name:getattr(settings,name) for name,_ in settings._fields_})
    finally:
        rc=dll.rpfem_accuracy_destroy(handle)
        deadline=time.monotonic()+10
        while rc==-8 and time.monotonic()<deadline:
            dll.rpfem_accuracy_cancel(handle);time.sleep(.02);rc=dll.rpfem_accuracy_destroy(handle)
        require(rc==0,'Worker cleanup failed')


def main():
    parser=argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--lp',required=True,type=Path);parser.add_argument('--dll',required=True,type=Path)
    parser.add_argument('--output',required=True,type=Path);parser.add_argument('--seconds',default=30.,type=float)
    args=parser.parse_args();require(0<args.seconds<=30,'Per solve budget must be <=30 seconds')
    require(not args.output.exists(),'Refusing to overwrite experiment output');args.output.mkdir(parents=True)
    before={str(p.resolve()):sha(p) for p in [args.lp,args.lp.with_suffix(args.lp.suffix+'.context.tsv'),args.dll,Path(__file__)]}
    for path in [args.lp,args.lp.with_suffix(args.lp.suffix+'.context.tsv'),args.dll]:(args.output/path.name).write_bytes(path.read_bytes())
    require(all(sha(p)==s for p,s in before.items()),'Source changed while snapshotting')
    lp=parse(args.output/args.lp.name);mesh,geometry=model(lp);prepared=core.Prepared(mesh,lp['fs'],'upper',8)
    dll,ns=abi((args.output/args.dll.name).resolve())
    report=dict(status='RUNNING',scope='Same original LP/DLL settings only; no numerical bound or root published',
                source_sha256=before,lp_sha256=sha(args.output/args.lp.name),fs=lp['fs'],q=8,
                dimensions={k:lp[k] for k in ['n','m','nz','neq']},geometry=geometry,trials=[])
    destination=args.output/'RESULTS.json'
    def save():core.write(destination,report)
    report['saved_export_physical_audit']=core.evaluate(prepared,lp['saved']);save()
    for equ,tol in [(0,1e-13),(1,1e-13),(0,1e-10),(1,1e-10)]:
        x,z,trial=solve(dll,ns,lp,equ,tol,args.seconds)
        trial.update(metrics(lp,x,z));trial['independent_original_q8']=core.evaluate(prepared,x)
        name=f'equ{equ}_tol{tol:g}.npz';np.savez_compressed(args.output/name,x=x,z=z)
        trial['vectors']=name;trial['vectors_sha256']=sha(args.output/name)
        report['trials'].append(trial);save()
        print(json.dumps({k:trial[k] for k in ['equilibrate','tolerance','wall_seconds','original_equality','original_primal_cone','literal_primal_gate']}),flush=True)
    report['source_sha256_after']={p:sha(p) for p in before}
    require(report['source_sha256_after']==before,'Source changed during comparison')
    report['status']='COMPLETED_DIAGNOSTIC_ONLY';save()


if __name__=='__main__':main()
