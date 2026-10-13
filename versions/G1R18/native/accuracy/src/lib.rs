//! Asynchronous owned Clarabel bridge. Original physical audits are external.
mod equality;
use clarabel::{algebra::CscMatrix, solver::*};
use std::{collections::HashMap, mem::{align_of,size_of}, panic::{catch_unwind,AssertUnwindSafe},
    ptr, slice, sync::{Arc,Mutex,MutexGuard,OnceLock,atomic::{AtomicBool,AtomicU64,AtomicUsize,Ordering}},
    thread::{self,JoinHandle}};

const ABI:u32=1;
const INVALID:i32=-1; const LIMIT:i32=-2; const HANDLE:i32=-3; const STATE:i32=-4;
const BUFFER:i32=-5; const PANIC:i32=-6; const SOLVER:i32=-7; const BUSY:i32=-8;
const CREATED:u32=0; const RUNNING:u32=1; const DONE:u32=2; const CANCELLED:u32=3; const ERROR:u32=4;
const MAX_N:usize=2_000_000; const MAX_M:usize=8_000_000; const MAX_NNZ:usize=50_000_000;
static NEXT:AtomicU64=AtomicU64::new(1);
static WORKERS:AtomicUsize=AtomicUsize::new(0);

#[repr(C)]
#[derive(Clone,Copy)]
pub struct RpaProblem { pub struct_size:u32,pub abi_version:u32,pub n:u32,pub m:u32,pub nnz:u32,pub cone_count:u32,
    pub row_ptr:*const i32,pub col_index:*const i32,pub a_values:*const f64,pub b:*const f64,
    pub pdiag:*const f64,pub q:*const f64,pub cone_kind:*const i32,pub cone_dim:*const i32 }
#[repr(C)]
#[derive(Clone,Copy)]
pub struct RpaSettings { pub struct_size:u32,pub abi_version:u32,pub max_iter:u32,pub max_threads:u32,
    pub time_limit:f64,pub tol_gap_abs:f64,pub tol_gap_rel:f64,pub tol_feas:f64,
    pub tol_infeas_abs:f64,pub tol_infeas_rel:f64,pub iterative_refinement_reltol:f64,pub iterative_refinement_abstol:f64,
    pub iterative_refinement_max_iter:u32,pub equilibrate:u32,pub static_regularization_constant:f64,
    pub dynamic_regularization_eps:f64,pub dynamic_regularization_delta:f64,
    pub static_regularization_enable:u32,pub dynamic_regularization_enable:u32,pub direct_method:u32,pub reserved:u32 }
#[repr(C)]
#[derive(Clone,Copy,Default)]
pub struct RpaProgress { pub struct_size:u32,pub abi_version:u32,pub state:u32,pub solver_status:u32,
    pub iterations:u32,pub cancel_requested:u32,pub n:u32,pub m:u32,
    pub solve_time:f64,pub primal_objective:f64,pub dual_objective:f64,pub primal_residual:f64,pub dual_residual:f64,
    pub diagnostics_valid:f64 }
struct OwnedProblem { n:usize,m:usize,a:CscMatrix<f64>,b:Vec<f64>,p:Vec<f64>,q:Vec<f64>,cones:Vec<SupportedConeT<f64>> }
struct Output { x:Vec<f64>,z:Vec<f64> }
struct JobState { progress:RpaProgress,output:Option<Output> }
struct Job { problem:Mutex<Option<OwnedProblem>>,settings:RpaSettings,cancel:AtomicBool,closed:AtomicBool,
    state:Mutex<JobState>,worker:Mutex<Option<JoinHandle<()>>> }
static REGISTRY:OnceLock<Mutex<HashMap<u64,Arc<Job>>>>=OnceLock::new();
fn lock<T>(m:&Mutex<T>)->MutexGuard<'_,T>{m.lock().unwrap_or_else(|e|e.into_inner())}
fn registry()->&'static Mutex<HashMap<u64,Arc<Job>>>{REGISTRY.get_or_init(||Mutex::new(HashMap::new()))}
fn job(h:u64)->Result<Arc<Job>,i32>{if h==0{return Err(HANDLE)} lock(registry()).get(&h).cloned().ok_or(HANDLE)}
fn ffi(f:impl FnOnce()->Result<(),i32>)->i32 {match catch_unwind(AssertUnwindSafe(f)){Ok(Ok(()))=>0,Ok(Err(e))=>e,Err(_)=>PANIC}}
fn checked_ptr<T>(p:*const T,len:usize)->Result<(),i32>{
    if len==0{return Ok(())} let start=p as usize;
    if p.is_null() || start%align_of::<T>()!=0 || len.checked_mul(size_of::<T>()).and_then(|v|start.checked_add(v)).is_none(){return Err(INVALID)}
    Ok(())
}
fn overlap<T,U>(a:*const T,na:usize,b:*const U,nb:usize)->bool {
    if na==0||nb==0{return false}
    let sa=a as usize;let sb=b as usize;
    let ea=sa.saturating_add(na.saturating_mul(size_of::<T>()));
    let eb=sb.saturating_add(nb.saturating_mul(size_of::<U>()));sa<eb&&sb<ea
}
unsafe fn copied<T:Copy>(p:*const T,len:usize)->Result<Vec<T>,i32>{checked_ptr(p,len)?;if len==0{Ok(Vec::new())}else{Ok(slice::from_raw_parts(p,len).to_vec())}}
unsafe fn header<T>(p:*const T)->Result<(),i32>{checked_ptr(p,1)?;let v=slice::from_raw_parts(p as *const u32,2);
    if v[0] as usize!=size_of::<T>() || v[1]!=ABI{return Err(INVALID)} Ok(())}
fn finite(v:&[f64])->bool{v.iter().all(|x|x.is_finite())}
fn settings(s:&RpaSettings)->Result<DefaultSettings<f64>,i32>{
    let positive=[s.time_limit,s.tol_gap_abs,s.tol_gap_rel,s.tol_feas,s.tol_infeas_abs,s.tol_infeas_rel,
        s.iterative_refinement_reltol,s.iterative_refinement_abstol,s.dynamic_regularization_eps,s.dynamic_regularization_delta];
    if !positive.iter().all(|v|v.is_finite()&&*v>0.) || s.time_limit>86400. || s.max_iter==0 || s.max_iter>100000
        || s.max_threads==0 || s.max_threads>8 || s.iterative_refinement_max_iter==0 || s.iterative_refinement_max_iter>1000
        || !s.static_regularization_constant.is_finite() || s.static_regularization_constant<0. || s.equilibrate>1
        || s.static_regularization_enable>1 || s.dynamic_regularization_enable>1 || s.direct_method>2 || s.reserved!=0{return Err(INVALID)}
    let mut d=DefaultSettings::<f64>::default();d.verbose=false;d.max_iter=s.max_iter;d.max_threads=s.max_threads;
    d.time_limit=s.time_limit;d.tol_gap_abs=s.tol_gap_abs;d.tol_gap_rel=s.tol_gap_rel;d.tol_feas=s.tol_feas;
    d.tol_infeas_abs=s.tol_infeas_abs;d.tol_infeas_rel=s.tol_infeas_rel;d.equilibrate_enable=s.equilibrate!=0;
    d.iterative_refinement_enable=true;d.iterative_refinement_max_iter=s.iterative_refinement_max_iter;
    d.iterative_refinement_reltol=s.iterative_refinement_reltol;d.iterative_refinement_abstol=s.iterative_refinement_abstol;
    d.static_regularization_enable=s.static_regularization_enable!=0;d.static_regularization_constant=s.static_regularization_constant;
    d.dynamic_regularization_enable=s.dynamic_regularization_enable!=0;d.dynamic_regularization_eps=s.dynamic_regularization_eps;
    d.dynamic_regularization_delta=s.dynamic_regularization_delta;
    d.direct_solve_method=match s.direct_method {0=>"auto",1=>"qdldl",_=>"faer"}.to_string();Ok(d)
}
unsafe fn problem(p:&RpaProblem)->Result<OwnedProblem,i32>{
    let(n,m,nz,k)=(p.n as usize,p.m as usize,p.nnz as usize,p.cone_count as usize);
    if n==0||n>MAX_N||m>MAX_M||nz>MAX_NNZ||k>m{return Err(LIMIT)}
    let rows=copied(p.row_ptr,m+1)?;let cols=copied(p.col_index,nz)?;let vals=copied(p.a_values,nz)?;
    let b=copied(p.b,m)?;let pd=copied(p.pdiag,n)?;let q=copied(p.q,n)?;
    let kinds=copied(p.cone_kind,k)?;let dims=copied(p.cone_dim,k)?;
    if !finite(&vals)||!finite(&b)||!finite(&pd)||!finite(&q)||pd.iter().any(|v|*v<0.)||rows[0]!=0||rows[m]!=nz as i32{return Err(INVALID)}
    let mut cones=Vec::with_capacity(k);let mut total=0usize;
    for (&kind,&dim) in kinds.iter().zip(&dims){if dim<=0{return Err(INVALID)} let dim=dim as usize;
        total=total.checked_add(dim).ok_or(LIMIT)?;
        cones.push(match kind {0=>ZeroConeT(dim),1=>NonnegativeConeT(dim),3 if dim==3=>SecondOrderConeT(dim),_=>return Err(INVALID)});
    } if total!=m{return Err(INVALID)}
    let mut cp=vec![0usize;n+1];
    for r in 0..m {let(a,e)=(rows[r],rows[r+1]);if a<0||e<a||e as usize>nz{return Err(INVALID)}
        let mut prev=None;
        for j in a as usize..e as usize {let c=cols[j];if c<0||c as usize>=n||prev.is_some_and(|v|v>=c){return Err(INVALID)}
            prev=Some(c);cp[c as usize+1]+=1;}
    }
    for c in 0..n {cp[c+1]+=cp[c];} let mut at=cp[..n].to_vec();let mut rv=vec![0usize;nz];let mut av=vec![0.;nz];
    for r in 0..m {for j in rows[r] as usize..rows[r+1] as usize{let c=cols[j] as usize;let i=at[c];rv[i]=r;av[i]=vals[j];at[c]+=1;}}
    Ok(OwnedProblem{n,m,a:CscMatrix::new(m,n,cp,rv,av),b,p:pd,q,cones})
}
fn diagnostics(p:&mut RpaProgress,t:f64,po:f64,do_:f64,rp:f64,rd:f64){
    let values=[t,po,do_,rp,rd];p.diagnostics_valid=if finite(&values){1.}else{0.};
    let clean=|v:f64|if v.is_finite(){v}else{0.};p.solve_time=clean(t);p.primal_objective=clean(po);
    p.dual_objective=clean(do_);p.primal_residual=clean(rp);p.dual_residual=clean(rd);
}
struct WorkerCount;
impl Drop for WorkerCount {fn drop(&mut self){WORKERS.fetch_sub(1,Ordering::AcqRel);}}
fn worker(j:Arc<Job>,data:OwnedProblem){
    let _count=WorkerCount;
    let solved=catch_unwind(AssertUnwindSafe(||->Result<Output,i32>{
        if j.cancel.load(Ordering::Acquire){return Err(STATE)}
        let p=CscMatrix::new(data.n,data.n,(0..=data.n).collect(),(0..data.n).collect(),data.p);
        let mut solver=DefaultSolver::new(&p,&data.q,&data.a,&data.b,&data.cones,settings(&j.settings)?).map_err(|_|SOLVER)?;
        let callback=j.clone();solver.set_termination_callback(move |info:&DefaultInfo<f64>|{
            let mut s=lock(&callback.state);s.progress.iterations=info.iterations;s.progress.solver_status=info.status as u32;
            diagnostics(&mut s.progress,info.solve_time,info.cost_primal,info.cost_dual,info.res_primal,info.res_dual);
            callback.cancel.load(Ordering::Acquire)
        });solver.solve();
        let sol=&solver.solution;if sol.x.len()!=data.n||sol.z.len()!=data.m||!finite(&sol.x)||!finite(&sol.z){return Err(SOLVER)}
        let mut s=lock(&j.state);s.progress.iterations=sol.iterations;s.progress.solver_status=sol.status as u32;
        diagnostics(&mut s.progress,sol.solve_time,sol.obj_val,sol.obj_val_dual,sol.r_prim,sol.r_dual);
        Ok(Output{x:sol.x.clone(),z:sol.z.clone()})
    }));
    let mut s=lock(&j.state);s.progress.cancel_requested=j.cancel.load(Ordering::Acquire) as u32;
    match solved {Ok(Ok(out))=>{s.output=Some(out);s.progress.state=if s.progress.cancel_requested!=0{CANCELLED}else{DONE}},
        _=>{s.progress.state=if s.progress.cancel_requested!=0{CANCELLED}else{ERROR};}}
}
#[no_mangle] pub extern "C" fn rpfem_accuracy_abi_version()->u32{ABI}
#[no_mangle] pub unsafe extern "C" fn rpfem_accuracy_default_settings(out:*mut RpaSettings)->i32{ffi(||{
    checked_ptr(out,1)?;let d=DefaultSettings::<f64>::default();ptr::write(out,RpaSettings{struct_size:size_of::<RpaSettings>() as u32,abi_version:ABI,
        max_iter:350,max_threads:1,time_limit:600.,tol_gap_abs:1e-13,tol_gap_rel:1e-13,tol_feas:1e-13,tol_infeas_abs:1e-13,tol_infeas_rel:1e-13,
        iterative_refinement_reltol:1e-15,iterative_refinement_abstol:1e-15,iterative_refinement_max_iter:30,equilibrate:0,
        static_regularization_constant:d.static_regularization_constant,dynamic_regularization_eps:d.dynamic_regularization_eps,
        dynamic_regularization_delta:d.dynamic_regularization_delta,static_regularization_enable:1,dynamic_regularization_enable:1,direct_method:0,reserved:0});Ok(())})}
#[no_mangle] pub unsafe extern "C" fn rpfem_accuracy_create(p:*const RpaProblem,s:*const RpaSettings,out:*mut u64)->i32{ffi(||{
    header(p)?;header(s)?;checked_ptr(out,1)?;let config=ptr::read(s);settings(&config)?;let data=problem(&ptr::read(p))?;
    let progress=RpaProgress{struct_size:size_of::<RpaProgress>() as u32,abi_version:ABI,state:CREATED,n:data.n as u32,m:data.m as u32,..Default::default()};
    let owned=Arc::new(Job{problem:Mutex::new(Some(data)),settings:config,cancel:AtomicBool::new(false),closed:AtomicBool::new(false),
        state:Mutex::new(JobState{progress,output:None}),worker:Mutex::new(None)});
    let mut map=lock(registry());if map.len()>=64{return Err(LIMIT)} let h=NEXT.fetch_add(1,Ordering::AcqRel);if h==0{return Err(LIMIT)}
    map.insert(h,owned);ptr::write(out,h);Ok(())})}
#[no_mangle] pub extern "C" fn rpfem_accuracy_start(h:u64)->i32{ffi(||{
    let j=job(h)?;let mut s=lock(&j.state);if j.closed.load(Ordering::Acquire)||s.progress.state!=CREATED{return Err(STATE)}
    if WORKERS.fetch_add(1,Ordering::AcqRel)>=8{WORKERS.fetch_sub(1,Ordering::AcqRel);return Err(BUSY)}
    let data=match lock(&j.problem).take(){Some(data)=>data,None=>{WORKERS.fetch_sub(1,Ordering::AcqRel);return Err(STATE)}};
    s.progress.state=RUNNING;let copy=j.clone();
    match thread::Builder::new().name("rpfem-accuracy".into()).spawn(move||worker(copy,data)){
        Ok(t)=>*lock(&j.worker)=Some(t),Err(_)=>{WORKERS.fetch_sub(1,Ordering::AcqRel);s.progress.state=ERROR;return Err(SOLVER)}} Ok(())})}
#[no_mangle] pub unsafe extern "C" fn rpfem_accuracy_poll(h:u64,out:*mut RpaProgress)->i32{ffi(||{
    header(out)?;let j=job(h)?;let mut p=lock(&j.state).progress;p.cancel_requested=j.cancel.load(Ordering::Acquire) as u32;ptr::write(out,p);Ok(())})}
#[no_mangle] pub extern "C" fn rpfem_accuracy_cancel(h:u64)->i32{ffi(||{let j=job(h)?;j.cancel.store(true,Ordering::Release);Ok(())})}
#[no_mangle] pub unsafe extern "C" fn rpfem_accuracy_result(h:u64,x:*mut f64,nx:u32,z:*mut f64,nz:u32,out:*mut RpaProgress)->i32{ffi(||{
    header(out)?;let j=job(h)?;let s=lock(&j.state);if s.progress.state!=DONE&&s.progress.state!=CANCELLED{return Err(STATE)}
    if nx!=s.progress.n||nz!=s.progress.m{return Err(BUFFER)} checked_ptr(x,nx as usize)?;checked_ptr(z,nz as usize)?;
    if overlap(x,nx as usize,z,nz as usize)||overlap(x,nx as usize,out,1)||overlap(z,nz as usize,out,1){return Err(INVALID)}
    let data=s.output.as_ref().ok_or(STATE)?;if !finite(&data.x)||!finite(&data.z){return Err(SOLVER)}
    if nx!=0{ptr::copy_nonoverlapping(data.x.as_ptr(),x,nx as usize)} if nz!=0{ptr::copy_nonoverlapping(data.z.as_ptr(),z,nz as usize)}
    let mut progress=s.progress;progress.cancel_requested=j.cancel.load(Ordering::Acquire) as u32;
    ptr::write(out,progress);Ok(())})}
#[no_mangle] pub extern "C" fn rpfem_accuracy_destroy(h:u64)->i32{ffi(||{
    let j=job(h)?;let s=lock(&j.state);if s.progress.state==RUNNING {j.cancel.store(true,Ordering::Release);return Err(BUSY)}
    let mut thread=lock(&j.worker);if thread.as_ref().is_some_and(|t|!t.is_finished()){return Err(BUSY)}
    j.closed.store(true,Ordering::Release);lock(registry()).remove(&h).ok_or(HANDLE)?;drop(s);
    if let Some(t)=thread.take(){t.join().map_err(|_|PANIC)?;}Ok(())})}

#[cfg(test)] mod tests {
    use super::*;
    #[test] fn abi_sizes(){assert_eq!(size_of::<RpaProblem>(),88);assert_eq!(size_of::<RpaSettings>(),128);assert_eq!(size_of::<RpaProgress>(),80);}
    #[test] fn rejects_unknown_handles(){assert_eq!(rpfem_accuracy_start(0),HANDLE);assert_eq!(rpfem_accuracy_destroy(u64::MAX),HANDLE);}
}
