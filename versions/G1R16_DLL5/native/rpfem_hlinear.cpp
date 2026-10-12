/* DLL4: resumable checked HSD solve/refinement. Owns all work vectors and
 * borrows fresh model/factor pointers only within each call. LGPL-2.1-or-later. */
#define NOMINMAX
#include <windows.h>
#include <cstdint>
#include <vector>
#include <unordered_map>
#include <new>
#include <cstring>
#include <algorithm>
struct DD {double hi,lo;};
extern "C" int __stdcall rpx_backsolve(int32_t,int32_t,int32_t,int32_t,const int32_t*,const int32_t*,const double*,const double*,const double*,double*);
extern "C" int __stdcall rpx_dd_backsolve(int32_t,int32_t,int32_t,int32_t,const int32_t*,const int32_t*,const DD*,const DD*,const double*,DD*);
extern "C" int __stdcall rpx_lu_retained_solve(uintptr_t,int32_t,int32_t,int32_t,int32_t,int32_t,const double*,double*,int32_t*);
extern "C" int __stdcall rpx_solver_vectors(int32_t,int32_t,int32_t,int32_t,const uintptr_t*,double,double*);
extern "C" int __stdcall rpx_hoperator(int32_t,int32_t,int32_t,int32_t,int32_t,int32_t,int32_t,int32_t,const uintptr_t*);
extern "C" double rpx_fp_add(double,double);
extern "C" double rpx_fp_mul(double,double);
extern "C" int32_t rpx_fp_greater_product(double,double,double);
#define API extern "C" __declspec(dllexport) int __stdcall
namespace {
struct Job {
 int32_t n,nx,nb,ns,kind=0,factor=0,stage=0,phase=0,cursor=0,j=0,checks=0,corrections=0,finite_checks=0,unshifted=0;uint64_t epoch=0;
 bool active=false;double best=1e300,error=0,previous=0,history[9]={};
 std::vector<double> scaled,solution,residual,correction;
 std::vector<DD> bdd,xdd,prod,bestdd,ddsolution,opx,opv,opg,oph;
 Job(int32_t n0,int32_t nx0,int32_t nb0,int32_t ns0):n(n0),nx(nx0),nb(nb0),ns(ns0),scaled(n),solution(n),residual(n),correction(n),bdd(n),xdd(n),prod(n),bestdd(n),ddsolution(n),opx(n),opv(n),opg(ns),oph(ns){}
};
std::unordered_map<uintptr_t,Job*> jobs;uintptr_t next_id=0;
struct Failure {int code;};
Job*get(uintptr_t h){auto i=jobs.find(h);if(i==jobs.end())throw Failure{318};return i->second;}
template<class F>int checked(F f){try{return f();}catch(Failure e){return e.code;}catch(const std::bad_alloc&){return 207;}catch(...){return 208;}}
uintptr_t ptr(const void*p){return reinterpret_cast<uintptr_t>(p);}
}
API rpx_hlinear_create(int32_t n,int32_t nx,int32_t nb,int32_t ns,uintptr_t *handle){if(!handle||n<1||n>1048576||nx<1||nx>n||nb<1||ns<1)return 100;*handle=0;
 return checked([&](){Job*p=new Job(n,nx,nb,ns);try{if(next_id==UINTPTR_MAX)throw Failure{318};uintptr_t id=++next_id;jobs.emplace(id,p);*handle=id;}catch(...){delete p;throw;}return 0;});}
API rpx_hlinear_destroy(uintptr_t handle){return checked([&](){auto p=get(handle);jobs.erase(handle);delete p;return 0;});}
API rpx_hlinear_begin(uintptr_t handle,uint64_t epoch,int32_t factor,int32_t kind,int32_t unshifted,const uintptr_t *args){if(!args||kind<0||kind>2||epoch==0)return 100;
 return checked([&](){auto p=get(handle);if(p->active)throw Failure{318};uintptr_t a[5]={args[0],args[1],args[2],ptr(p->bdd.data()),ptr(p->scaled.data())};double m[1]={};
  int code=rpx_solver_vectors(2,p->n,0,p->n,a,0,m);if(code)throw Failure{code};
  p->epoch=epoch;p->factor=factor;p->kind=kind;p->unshifted=unshifted;p->stage=0;p->phase=0;p->cursor=0;p->j=0;p->checks=0;p->corrections=0;p->finite_checks=0;
  p->best=1e300;p->error=0;p->previous=0;std::fill(p->history,p->history+9,0.0);p->active=true;return 0;});}
API rpx_hlinear_step(uintptr_t handle,uint64_t epoch,int32_t factor,const uintptr_t *opargs,const uintptr_t *fargs,uintptr_t lu,int32_t ms,int32_t *done){
 if(!opargs||!fargs||!done||ms<1||ms>1000)return 100;*done=0;
 return checked([&](){auto p=get(handle);if(!p->active||p->epoch!=epoch||p->factor!=factor)throw Failure{318};LARGE_INTEGER freq,clock;QueryPerformanceFrequency(&freq);QueryPerformanceCounter(&clock);LONGLONG deadline=clock.QuadPart+freq.QuadPart*ms/1000;
  do {
   int code=0;int32_t total=p->n,first=p->cursor,last=std::min(total,first+32);double m[1]={p->error};
   if(p->stage==0||p->stage==4){ /* initial or correction triangular solve */
    bool initial=p->stage==0;const double *rhs=initial?p->scaled.data():p->residual.data();double *out=initial?p->solution.data():p->correction.data();
    if((p->kind==2&&p->phase==1)||(p->kind!=2&&p->phase==3)){last=total-first;first=std::max(0,last-32);}
    if(p->kind==2){code=rpx_lu_retained_solve(lu,p->factor,p->n,p->phase,first,last,rhs,out,&p->finite_checks);}
    else if(p->kind==1){code=rpx_dd_backsolve(p->n,p->phase,first,last,reinterpret_cast<const int32_t*>(fargs[0]),reinterpret_cast<const int32_t*>(fargs[1]),reinterpret_cast<const DD*>(fargs[2]),reinterpret_cast<const DD*>(fargs[3]),rhs,p->ddsolution.data());}
    else code=rpx_backsolve(p->n,p->phase,first,last,reinterpret_cast<const int32_t*>(fargs[0]),reinterpret_cast<const int32_t*>(fargs[1]),reinterpret_cast<const double*>(fargs[2]),reinterpret_cast<const double*>(fargs[3]),rhs,out);
   }else if(p->stage==1||p->stage==5){ /* DD iterate initialization/correction */
    uintptr_t a[3]={ptr(p->xdd.data()),ptr(p->stage==1?p->solution.data():p->correction.data()),ptr(p->ddsolution.data())};
    code=rpx_solver_vectors(p->stage==1?4:3,p->n,first,last,a,p->kind==1?1.0:0.0,m);
   }else if(p->stage==2){ /* checked original DD operator */
    uintptr_t a[20];std::copy(opargs,opargs+14,a);a[14]=ptr(p->xdd.data());a[15]=ptr(p->prod.data());a[16]=ptr(p->opx.data());a[17]=ptr(p->opv.data());a[18]=ptr(p->opg.data());a[19]=ptr(p->oph.data());
    if(p->phase==0){code=rpx_hoperator(p->n,p->nx,p->nb,p->ns,-1,0,0,p->unshifted,a);p->cursor=total;}
    else{int32_t phase=p->phase-1;if(phase==1)total=p->nb;if(phase==2)total=p->n-p->nx;last=std::min(total,first+32);
     if(total>0)code=rpx_hoperator(p->n,p->nx,p->nb,p->ns,phase,first,last,p->unshifted,a);}
   }else if(p->stage==3){uintptr_t a[4]={ptr(p->bdd.data()),ptr(p->prod.data()),opargs[1],ptr(p->residual.data())};
    /* opargs[1] is DD coefficient data; plain Double scales are fargs[4]. */
    a[2]=fargs[4];code=rpx_solver_vectors(1,p->n,first,last,a,0,m);p->error=m[0];
   }else throw Failure{318};
   if(code){p->active=false;throw Failure{code};}
   p->cursor=std::min(total,p->cursor+32);
   if(p->cursor==total){p->cursor=0;
    if(p->stage==0||p->stage==4){p->phase++;int32_t phases=p->kind==2?3:4;
     if(p->phase==phases){p->phase=0;p->stage=p->stage==0?1:5;}}
    else if(p->stage==1||p->stage==5){if(p->stage==5){p->corrections++;p->j++;}p->stage=2;p->phase=0;p->error=0;}
    else if(p->stage==2){p->phase++;if(p->phase==5){p->stage=3;p->phase=0;}}
    else if(p->stage==3){p->history[p->j]=p->error;p->checks++;if(p->error<p->best){p->best=p->error;p->bestdd=p->xdd;}
     if(p->error<=1e-11||(p->j>0&&rpx_fp_greater_product(p->error,p->previous,10))||(p->j>=3&&(p->kind==2||rpx_fp_greater_product(p->error,p->previous,.25)))||p->j==8){p->active=false;*done=1;return 0;}
     p->previous=p->error;p->stage=4;p->phase=0;
    }
   }
   QueryPerformanceCounter(&clock);
  }while(clock.QuadPart<deadline);return 0;
 });
}
API rpx_hlinear_result(uintptr_t handle,uint64_t epoch,DD *best,double *scaled,DD *bdd,double *metrics){if(!best||!scaled||!bdd||!metrics)return 100;
 return checked([&](){auto p=get(handle);if(p->active||p->epoch!=epoch||p->checks<1)throw Failure{318};
  std::memcpy(best,p->bestdd.data(),size_t(p->n)*16);std::memcpy(scaled,p->scaled.data(),size_t(p->n)*8);std::memcpy(bdd,p->bdd.data(),size_t(p->n)*16);
  metrics[0]=p->checks;metrics[1]=p->corrections;metrics[2]=p->best;metrics[3]=p->error;for(int i=0;i<9;i++)metrics[4+i]=p->history[i];metrics[13]=p->finite_checks;return 0;});}
API rpx_hlinear_live_handles(void){return static_cast<int>(jobs.size());}
