"""Both unchanged native ABIs in one PE: meaningful small regressions only."""
from pathlib import Path
import argparse
import ctypes as C
import hashlib
import importlib.util
import json
import os
import shutil
import struct
import subprocess
import sys
import time
import numpy as np

HERE=Path(__file__).resolve().parents[1]
REPO=HERE.parents[3]
LEGACY=REPO/'versions/G1R17'
ACCURACY_EXPORTS={'rpfem_accuracy_'+x for x in ('abi_version','default_settings','create','start','poll','cancel','result','destroy','equalities')}


def sha(p):return hashlib.sha256(Path(p).read_bytes()).hexdigest()


def write(p,value):Path(p).write_text(json.dumps(value,indent=2,allow_nan=False)+'\n',encoding='utf8')


def pe_info(path):
    data=Path(path).read_bytes();pe=struct.unpack_from('<I',data,0x3c)[0]
    assert data[pe:pe+4]==b'PE\0\0' and struct.unpack_from('<H',data,pe+4)[0]==0x8664
    count=struct.unpack_from('<H',data,pe+6)[0];size=struct.unpack_from('<H',data,pe+20)[0];opt=pe+24
    assert struct.unpack_from('<H',data,opt)[0]==0x20b
    sections=[]
    for k in range(count):
        s=opt+size+40*k;vs,va,rs,raw=struct.unpack_from('<IIII',data,s+8)
        sections.append((va,max(vs,rs),raw))
    def offset(rva):
        for va,length,raw in sections:
            if va<=rva<va+length:return raw+rva-va
        raise AssertionError('RVA outside PE sections')
    def string(rva):
        at=offset(rva);return data[at:data.index(b'\0',at)].decode('ascii')
    export_rva=struct.unpack_from('<I',data,opt+112)[0];entry=offset(export_rva)
    n=struct.unpack_from('<I',data,entry+24)[0];names=offset(struct.unpack_from('<I',data,entry+32)[0])
    exports={string(struct.unpack_from('<I',data,names+4*k)[0]) for k in range(n)}
    imports=[];rva=struct.unpack_from('<I',data,opt+120)[0]
    if rva:
        at=offset(rva)
        while True:
            descriptor=struct.unpack_from('<5I',data,at)
            if not any(descriptor):break
            imports.append(string(descriptor[3]));at+=20
    return {'machine':'AMD64','exports':sorted(exports),'imports':sorted(imports)}


class DD(C.Structure):_fields_=[('hi',C.c_double),('lo',C.c_double)]


def factor_regression(original,unified):
    i,p=C.c_int32,C.c_void_p
    for dll in (original,unified):
        dll.rpx_numeric.argtypes=[i,i,i]+[p]*13;dll.rpx_numeric.restype=i
        dll.rpx_dd_numeric.argtypes=[i,i,i]+[p]*13;dll.rpx_dd_numeric.restype=i
        dll.rpx_backsolve.argtypes=[i,i,i,i]+[p]*6;dll.rpx_backsolve.restype=i
        dll.rpx_dd_backsolve.argtypes=[i,i,i,i]+[p]*6;dll.rpx_dd_backsolve.restype=i
        dll.rpx_dd_test.argtypes=[i,C.POINTER(DD),C.POINTER(DD),C.POINTER(DD)];dll.rpx_dd_test.restype=i
    ptr=lambda a:p(a.ctypes.data)
    ap=np.array([0,1,3,6],np.int32);ai=np.array([0,0,1,0,1,2],np.int32)
    cp=np.array([0,2,3,3],np.int32);ri=np.array([1,2,2],np.int32)
    rp=np.array([0,0,1,3],np.int32);rn=np.array([0,0,1],np.int32);wa=np.array([0,1,2],np.int32)
    rng=np.random.default_rng(190215);records=[]
    for k in range(60):
        B=rng.normal(size=(3,3));A=B@B.T+np.eye(3)*.5
        if k%2:A[1,1]*=-1
        av=np.array([A[0,0],A[0,1],A[1,1],A[0,2],A[1,2],A[2,2]])
        rhs=rng.normal(size=3)
        for dd in (False,True):
            values=np.zeros((6,2)) if dd else av.copy()
            if dd:values[:,0]=av
            outputs=[]
            for dll in (original,unified):
                shape=(3,2) if dd else (3,)
                l=np.zeros(shape);d=np.zeros(shape);y=np.zeros(shape);x=np.zeros(shape)
                minimum=np.array([1e300]);failure=np.array([-1],np.int32)
                f=dll.rpx_dd_numeric if dd else dll.rpx_numeric
                rc=f(3,0,3,*map(ptr,(ap,ai,values,cp,ri,rp,rn,wa,l,d,y,minimum,failure)))
                assert rc==0,rc
                back=dll.rpx_dd_backsolve if dd else dll.rpx_backsolve
                for phase in range(4):assert back(3,phase,0,3,*map(ptr,(cp,ri,l,d,rhs,x)))==0
                restored=x.sum(axis=1) if dd else x
                assert np.max(abs(A@restored-rhs))<2e-12
                outputs.append((l.tobytes(),d.tobytes(),x.tobytes(),minimum.tobytes(),failure.tobytes()))
            assert outputs[0]==outputs[1]
            records.append(dict(case=k,double_double=dd,bitwise_exact=True))
    errors=[]
    for operation,a,b in [(0,DD(1.,1e-17),DD(.1,-1e-19)),(1,DD(1.,1e-17),DD(.1,-1e-19)),
                          (2,DD(1.1,1e-17),DD(.3,1e-18)),(3,DD(1.1,1e-17),DD(.3,1e-18)),
                          (3,DD(1.,0.),DD(0.,0.)),(2,DD(1e-200,0.),DD(1e-200,0.))]:
        output=[]
        for dll in (original,unified):
            value=DD();code=dll.rpx_dd_test(operation,C.byref(a),C.byref(b),C.byref(value))
            output.append((code,bytes(value)))
        assert output[0]==output[1]
        errors.append(dict(operation=operation,code=output[0][0],bitwise_exact=True))
    return dict(dense_spd_and_indefinite_factor_backsolve_cases=len(records),
                bitwise_exact=True,cases=records,double_double_arithmetic_and_errors=errors)


def load_module(name,path):
    spec=importlib.util.spec_from_file_location(name,path);module=importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module);return module


def main(args):
    dll=Path(args.dll).resolve();baseline=LEGACY/'delivery/rpfem_native_x64.dll'
    folder=Path(args.output).resolve();assert not folder.exists();folder.mkdir(parents=True)
    protected={str(p):sha(p) for p in (dll,baseline,HERE/'baseline_source_sha256.json')}
    baseline_hashes=json.loads((HERE/'baseline_source_sha256.json').read_text())
    assert all(sha(REPO/p)==h for p,h in baseline_hashes.items())
    before=pe_info(baseline);after=pe_info(dll)
    oldset,newset=set(before['exports']),set(after['exports'])
    assert len(oldset)==94 and oldset.isdisjoint(ACCURACY_EXPORTS)
    assert newset==oldset|ACCURACY_EXPORTS and len(newset)==103
    assert not any('rpfem' in x.lower() or 'vcruntime' in x.lower() or 'msvcp' in x.lower() or
                   'ucrtbase' in x.lower() for x in after['imports'])
    isolated=folder/'legacy';(isolated/'delivery').mkdir(parents=True);(isolated/'results').mkdir()
    shutil.copyfile(dll,isolated/'delivery/rpfem_native_x64.dll')
    original=C.WinDLL(str(baseline));unified=C.WinDLL(str(dll))
    assert original.rpx_abi()==unified.rpx_abi()==0x00170005
    unified.rpfem_accuracy_abi_version.restype=C.c_uint32
    assert unified.rpfem_accuracy_abi_version()==1
    original.rpx_test_controlword.argtypes=[C.c_int32,C.c_int32]
    oldcw=original.rpx_test_controlword(0,0);original.rpx_test_controlword(1,0x137f)
    try:numeric=factor_regression(original,unified)
    finally:original.rpx_test_controlword(1,oldcw)
    sys.path.insert(0,str(LEGACY/'tools'));sys.path.insert(0,str(LEGACY/'python'))
    compiler=load_module('unified_compile_oracle',LEGACY/'tools/test_compile_python.py')
    compiler.ROOT=isolated;compiler.run()
    fan=load_module('unified_fan_oracle',LEGACY/'python/test_meshfan_native.py')
    fan.ROOT=isolated;fan.main()
    env=os.environ.copy();env.update(RPFEM_ACCURACY_DLL=str(dll),
        RPFEM_ACCURACY_QA_REPORT=str(folder/'accuracy_ffi.json'))
    command=[sys.executable,'-X','utf8','-m','unittest','discover','-s',str(REPO/'versions/G1R18/tests'),
             '-p','test_accuracy_native_ffi.py','-v']
    completed=subprocess.run(command,env=env,cwd=REPO,stdout=subprocess.PIPE,stderr=subprocess.STDOUT)
    (folder/'accuracy_ffi.log').write_bytes(completed.stdout)
    print(completed.stdout.decode('utf8',errors='replace'),flush=True)
    assert completed.returncode==0
    assert protected=={p:sha(p) for p in protected}
    assert all(sha(REPO/p)==h for p,h in baseline_hashes.items())
    report=dict(pass_all=True,dll=str(dll),dll_sha256=sha(dll),legacy_abi='0x00170005',accuracy_abi=1,
        legacy_exports=94,accuracy_exports=9,total_exports=103,pe=after,numeric=numeric,
        compiler_cases=120,fan_regression='actual geometry/guards/100 cancelled handles',
        accuracy_ffi_unittest_exit_code=completed.returncode,original_artifacts_unchanged=True,
        protected_sha256=protected,scope='small ABI regression; no new accuracy experiment or bound adoption')
    write(folder/'summary.json',report);print(json.dumps({k:report[k] for k in
        ('pass_all','dll_sha256','total_exports','compiler_cases','original_artifacts_unchanged')},indent=2))


if __name__=='__main__':
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--dll',type=Path,default=HERE/'bin/rpfem_native_x64.dll')
    p.add_argument('--output',type=Path,default=HERE/'validation'/('run_'+str(time.time_ns())))
    main(p.parse_args())
